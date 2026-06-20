"""Tests de la selección y los backends de motores de maquetación.

Se monkeypatchea la detección/ejecución para no depender de que Word o
LibreOffice estén instalados en la máquina de CI.
"""
import os

import pytest

from docx2pdf_py import converter as C
from docx2pdf_py import engines as E
from tests.conftest import document


@pytest.fixture
def docx(make_docx):
    """Un .docx mínimo real (convert() valida que la entrada sea un ZIP OOXML)."""
    return make_docx(document("<w:p><w:r><w:t>x</w:t></w:r></w:p>"))


# -- detección de LibreOffice ----------------------------------------------
def test_find_libreoffice_env_override(tmp_path, monkeypatch):
    fake = tmp_path / "soffice"
    fake.write_text("")
    monkeypatch.setenv("SOFFICE_BIN", str(fake))
    assert E.find_libreoffice() == str(fake)


def test_find_libreoffice_from_path(monkeypatch):
    monkeypatch.delenv("SOFFICE_BIN", raising=False)
    monkeypatch.setattr(E.shutil, "which",
                        lambda n: "/usr/bin/soffice" if n == "soffice" else None)
    assert E.find_libreoffice() == "/usr/bin/soffice"


def test_find_libreoffice_absent(monkeypatch):
    monkeypatch.delenv("SOFFICE_BIN", raising=False)
    monkeypatch.setattr(E.shutil, "which", lambda n: None)
    monkeypatch.setattr(E, "_SOFFICE_PATHS", ())
    assert E.find_libreoffice() is None


# -- selección de motor en convert() ----------------------------------------
def test_auto_falls_back_to_weasyprint(docx, monkeypatch):
    monkeypatch.setattr(E, "word_available", lambda: False)
    monkeypatch.setattr(E, "find_libreoffice", lambda: None)
    monkeypatch.setattr(C, "_convert_weasyprint",
                        lambda i, o: f"weasy:{o}")
    assert C.convert(docx, "out.pdf") == "weasy:out.pdf"


def test_auto_prefers_word_then_libreoffice(docx, monkeypatch):
    calls = []
    monkeypatch.setattr(E, "word_available", lambda: True)
    monkeypatch.setattr(E, "convert_word",
                        lambda i, o: calls.append("word") or f"word:{o}")
    monkeypatch.setattr(E, "find_libreoffice", lambda: "/usr/bin/soffice")
    assert C.convert(docx, "out.pdf") == "word:out.pdf"
    assert calls == ["word"]


def test_auto_degrades_when_engine_raises(docx, monkeypatch):
    monkeypatch.setattr(E, "word_available", lambda: True)

    def boom(i, o):
        raise RuntimeError("Word reventó")

    monkeypatch.setattr(E, "convert_word", boom)
    monkeypatch.setattr(E, "find_libreoffice", lambda: None)
    monkeypatch.setattr(C, "_convert_weasyprint", lambda i, o: f"weasy:{o}")
    assert C.convert(docx, "out.pdf") == "weasy:out.pdf"


def test_explicit_libreoffice_unavailable_raises(docx, monkeypatch):
    monkeypatch.setattr(E, "find_libreoffice", lambda: None)
    with pytest.raises(RuntimeError, match="LibreOffice"):
        C.convert(docx, "out.pdf", engine="libreoffice")


def test_explicit_word_unavailable_raises(docx, monkeypatch):
    monkeypatch.setattr(E, "word_available", lambda: False)
    with pytest.raises(RuntimeError, match="Word"):
        C.convert(docx, "out.pdf", engine="word")


def test_unknown_engine_raises():
    with pytest.raises(ValueError, match="motor desconocido"):
        C.convert("in.docx", "out.pdf", engine="ghostscript")


def test_default_engine_reports_availability(monkeypatch):
    monkeypatch.setattr(E, "word_available", lambda: False)
    monkeypatch.setattr(E, "find_libreoffice", lambda: "/usr/bin/soffice")
    assert E.default_engine() == "libreoffice"
    monkeypatch.setattr(E, "find_libreoffice", lambda: None)
    assert E.default_engine() == "weasyprint"


# -- backend de LibreOffice (subprocess simulado) ---------------------------
def test_convert_libreoffice_moves_output(tmp_path, monkeypatch):
    src = tmp_path / "doc.docx"
    src.write_bytes(b"PK\x03\x04")  # contenido irrelevante: run está simulado
    out = tmp_path / "sub" / "doc.pdf"

    def fake_run(cmd, capture_output, timeout):
        outdir = cmd[cmd.index("--outdir") + 1]
        stem = os.path.splitext(os.path.basename(cmd[-1]))[0]
        with open(os.path.join(outdir, stem + ".pdf"), "wb") as fh:
            fh.write(b"%PDF-1.4 generado")

        class R:
            returncode = 0
            stdout = b""
            stderr = b""

        return R()

    monkeypatch.setattr(E.subprocess, "run", fake_run)
    result = E.convert_libreoffice(str(src), str(out), soffice="/usr/bin/soffice")
    assert result == str(out)
    assert out.read_bytes().startswith(b"%PDF")


def test_convert_libreoffice_reports_failure(tmp_path, monkeypatch):
    src = tmp_path / "doc.docx"
    src.write_bytes(b"PK\x03\x04")

    def fake_run(cmd, capture_output, timeout):
        class R:
            returncode = 1
            stdout = b""
            stderr = b"Error: source file could not be loaded"

        return R()

    monkeypatch.setattr(E.subprocess, "run", fake_run)
    with pytest.raises(RuntimeError, match="no pudo convertir"):
        E.convert_libreoffice(str(src), str(tmp_path / "o.pdf"),
                              soffice="/usr/bin/soffice")


def test_convert_libreoffice_no_binary(monkeypatch):
    monkeypatch.setattr(E, "find_libreoffice", lambda: None)
    with pytest.raises(RuntimeError, match="no está disponible"):
        E.convert_libreoffice("in.docx", "out.pdf", soffice=None)
