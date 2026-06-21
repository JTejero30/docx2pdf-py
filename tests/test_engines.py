"""Tests de la selección y los backends de motores de maquetación.

Se monkeypatchea la detección/ejecución para no depender de que Word o
LibreOffice estén instalados en la máquina de CI.
"""
import os

import pytest

from docx2pdf_py import converter as C
from docx2pdf_py import engines as E
from tests.conftest import FAKE_PDF, document


def _write_pdf(path, prefix=""):
    with open(path, "wb") as stream:
        stream.write(FAKE_PDF)
    return str(path)


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
                        lambda i, o, options=None: _write_pdf(o))
    assert C.convert(docx, "out.pdf") == "out.pdf"


def test_auto_prefers_word_then_libreoffice(docx, monkeypatch):
    calls = []
    monkeypatch.setattr(E, "word_available", lambda: True)
    monkeypatch.setattr(E, "convert_word",
                        lambda i, o, timeout=120: calls.append("word") or _write_pdf(o))
    monkeypatch.setattr(E, "find_libreoffice", lambda: "/usr/bin/soffice")
    assert C.convert(docx, "out.pdf") == "out.pdf"
    assert calls == ["word"]


def test_auto_degrades_when_engine_raises(docx, monkeypatch):
    monkeypatch.setattr(E, "word_available", lambda: True)

    def boom(i, o, timeout=120):
        raise RuntimeError("Word reventó")

    monkeypatch.setattr(E, "convert_word", boom)
    monkeypatch.setattr(E, "find_libreoffice", lambda: None)
    monkeypatch.setattr(C, "_convert_weasyprint",
                        lambda i, o, options=None: _write_pdf(o))
    assert C.convert(docx, "out.pdf") == "out.pdf"


def test_explicit_libreoffice_unavailable_raises(docx, monkeypatch):
    monkeypatch.setattr(E, "find_libreoffice", lambda: None)
    with pytest.raises(RuntimeError, match="LibreOffice"):
        C.convert(docx, "out.pdf", engine="libreoffice")


def test_explicit_word_unavailable_raises(docx, monkeypatch):
    monkeypatch.setattr(E, "word_available", lambda: False)
    with pytest.raises(RuntimeError, match="Word"):
        C.convert(docx, "out.pdf", engine="word")


def test_unknown_engine_raises():
    with pytest.raises(ValueError, match="unknown engine"):
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
            fh.write(FAKE_PDF)

        class R:
            returncode = 0
            stdout = b""
            stderr = b""

        return R()

    monkeypatch.setattr(E, "run_process", lambda cmd, timeout: fake_run(cmd, True, timeout))
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

    monkeypatch.setattr(E, "run_process", lambda cmd, timeout: fake_run(cmd, True, timeout))
    with pytest.raises(RuntimeError, match="failed to convert"):
        E.convert_libreoffice(str(src), str(tmp_path / "o.pdf"),
                              soffice="/usr/bin/soffice")


def test_convert_libreoffice_no_binary(monkeypatch):
    monkeypatch.setattr(E, "find_libreoffice", lambda: None)
    with pytest.raises(RuntimeError, match="not available"):
        E.convert_libreoffice("in.docx", "out.pdf", soffice=None)


# -- entry-point engine discovery -------------------------------------------
def test_load_engine_registry_includes_builtins():
    from docx2pdf_py.backends import BUILTIN_ENGINES, load_engine_registry

    registry = load_engine_registry()
    builtin_names = {e.name for e in BUILTIN_ENGINES}
    registry_names = {e.name for e in registry}
    assert builtin_names <= registry_names


def test_load_engine_registry_skips_builtin_entry_points(monkeypatch):
    """Re-registering a built-in engine name via entry point must be a no-op."""
    import docx2pdf_py.backends as B

    class _FakeEP:
        name = "weasyprint"

        def load(self):
            raise AssertionError("should not be loaded")

    monkeypatch.setattr(B, "entry_points", lambda group: [_FakeEP()])
    registry = B.load_engine_registry()
    assert len(registry) == len(B.BUILTIN_ENGINES)


def test_load_engine_registry_logs_broken_entry_point(monkeypatch, caplog):
    import logging

    import docx2pdf_py.backends as B

    class _BadEP:
        name = "broken-engine"

        def load(self):
            raise ImportError("missing dep")

    monkeypatch.setattr(B, "entry_points", lambda group: [_BadEP()])
    with caplog.at_level(logging.WARNING, logger="docx2pdf_py.backends"):
        registry = B.load_engine_registry()
    assert len(registry) == len(B.BUILTIN_ENGINES)
    assert any("broken-engine" in r.message for r in caplog.records)
