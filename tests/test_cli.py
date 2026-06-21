"""Command-line behavior and user-facing reporting."""

import zipfile
from pathlib import Path

import pytest

from docx2pdf_py import ConversionResult, cli
from docx2pdf_py import converter as C
from tests.conftest import FAKE_PDF, document


def test_cli_reports_engine_actually_used(tmp_path, monkeypatch, capsys):
    source = tmp_path / "input.docx"
    source.write_bytes(b"placeholder")
    output = tmp_path / "output.pdf"
    monkeypatch.setattr(
        cli,
        "convert_detailed",
        lambda *args, **kwargs: ConversionResult(str(output), "weasyprint"),
    )

    assert cli.main([str(source), str(output)]) == 0
    assert "engine: weasyprint" in capsys.readouterr().out


def test_cli_refuses_to_overwrite_without_force(tmp_path):
    source = tmp_path / "input.docx"
    output = tmp_path / "output.pdf"
    source.write_bytes(b"placeholder")
    output.write_bytes(b"existing")

    with pytest.raises(SystemExit) as exc:
        cli.main([str(source), str(output)])
    assert exc.value.code == 2


def test_cli_quiet_suppresses_success_output(tmp_path, monkeypatch, capsys):
    source = tmp_path / "input.docx"
    source.write_bytes(b"placeholder")
    output = tmp_path / "output.pdf"
    monkeypatch.setattr(
        cli,
        "convert_detailed",
        lambda *args, **kwargs: ConversionResult(
            str(output), "libreoffice", ("word failed",)
        ),
    )

    assert cli.main([str(source), str(output), "--quiet"]) == 0
    captured = capsys.readouterr()
    assert captured.out == ""
    assert captured.err == ""


def _make_docx(path: Path) -> Path:
    """Write a minimal valid .docx at the given path."""
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("word/document.xml", document("<w:p/>"))
        z.writestr(
            "word/styles.xml",
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
        )
    return path


def test_cli_batch_mode_converts_multiple_files(tmp_path, monkeypatch, capsys):
    src1 = _make_docx(tmp_path / "a.docx")
    src2 = _make_docx(tmp_path / "b.docx")
    out_dir = tmp_path / "pdfs"

    def render(_input, output, options=None):
        with open(output, "wb") as stream:
            stream.write(FAKE_PDF)
        return str(output)

    monkeypatch.setattr(C, "_convert_weasyprint", render)
    rc = cli.main([
        str(src1), str(src2),
        "--output-dir", str(out_dir),
        "--engine", "weasyprint",
    ])

    assert rc == 0
    assert (out_dir / "a.pdf").exists()
    assert (out_dir / "b.pdf").exists()
    out = capsys.readouterr().out
    assert "OK" in out and "a.pdf" in out


def test_cli_batch_mode_reports_failure_and_returns_1(tmp_path, monkeypatch, capsys):
    src = _make_docx(tmp_path / "bad.docx")
    out_dir = tmp_path / "pdfs"

    monkeypatch.setattr(C, "_convert_weasyprint",
                        lambda *a, **kw: (_ for _ in ()).throw(RuntimeError("boom")))

    rc = cli.main([str(src), "--output-dir", str(out_dir), "--engine", "weasyprint"])

    assert rc == 1
    assert "ERR" in capsys.readouterr().err


def test_cli_batch_quiet_suppresses_output(tmp_path, monkeypatch, capsys):
    src = _make_docx(tmp_path / "c.docx")
    out_dir = tmp_path / "pdfs"

    def render(_input, output, options=None):
        with open(output, "wb") as stream:
            stream.write(FAKE_PDF)
        return str(output)

    monkeypatch.setattr(C, "_convert_weasyprint", render)
    rc = cli.main([
        str(src),
        "--output-dir", str(out_dir),
        "--engine", "weasyprint",
        "--quiet",
    ])

    assert rc == 0
    captured = capsys.readouterr()
    assert captured.out == ""
    assert captured.err == ""


def test_cli_too_many_positional_args_errors(tmp_path):
    with pytest.raises(SystemExit) as exc:
        cli.main(["a.docx", "b.docx", "c.docx"])
    assert exc.value.code == 2


def test_cli_verbose_single_file(tmp_path, monkeypatch, capsys):
    source = tmp_path / "input.docx"
    source.write_bytes(b"placeholder")
    output = tmp_path / "out.pdf"
    monkeypatch.setattr(
        cli,
        "convert_detailed",
        lambda *args, **kwargs: ConversionResult(str(output), "weasyprint"),
    )

    assert cli.main([str(source), str(output), "--verbose"]) == 0
    out = capsys.readouterr().out
    assert "Elapsed" in out


def test_cli_single_file_one_arg_defaults_output(tmp_path, monkeypatch, capsys):
    source = tmp_path / "input.docx"
    source.write_bytes(b"placeholder")
    output = tmp_path / "output.pdf"
    monkeypatch.setattr(
        cli,
        "convert_detailed",
        lambda *args, **kwargs: ConversionResult(str(output), "weasyprint"),
    )

    assert cli.main([str(source), "--force"]) == 0
    assert "engine: weasyprint" in capsys.readouterr().out


def test_cli_single_file_auto_discover(tmp_path, monkeypatch, capsys):
    source = tmp_path / "auto.docx"
    source.write_bytes(b"placeholder")
    output = tmp_path / "output.pdf"
    monkeypatch.setattr(
        cli,
        "convert_detailed",
        lambda *args, **kwargs: ConversionResult(str(output), "weasyprint"),
    )
    monkeypatch.setattr(cli.glob, "glob", lambda pattern: [str(source)])

    assert cli.main(["--force"]) == 0
    err = capsys.readouterr().err
    assert "auto-selected" in err


def test_cli_single_file_auto_discover_no_files(monkeypatch):
    monkeypatch.setattr(cli.glob, "glob", lambda pattern: [])
    with pytest.raises(SystemExit) as exc:
        cli.main([])
    assert exc.value.code == 2


def test_cli_fallback_flag_passed_to_options(tmp_path, monkeypatch, capsys):
    source = tmp_path / "input.docx"
    source.write_bytes(b"placeholder")
    output = tmp_path / "out.pdf"
    captured_options = {}

    def fake_convert(src, out, engine="auto", options=None):
        captured_options["fallback"] = options.fallback if options else None
        return ConversionResult(str(output), "weasyprint")

    monkeypatch.setattr(cli, "convert_detailed", fake_convert)
    assert cli.main([str(source), str(output), "--fallback", "never"]) == 0
    assert captured_options["fallback"] == "never"


def test_cli_batch_verbose_mode(tmp_path, monkeypatch, capsys):
    src = _make_docx(tmp_path / "v.docx")
    out_dir = tmp_path / "pdfs"

    def render(_input, output, options=None):
        with open(output, "wb") as stream:
            stream.write(FAKE_PDF)
        return str(output)

    monkeypatch.setattr(C, "_convert_weasyprint", render)
    rc = cli.main([
        str(src),
        "--output-dir", str(out_dir),
        "--engine", "weasyprint",
        "--verbose",
    ])

    assert rc == 0
    out = capsys.readouterr().out
    assert "engine:" in out


def test_cli_batch_auto_discover(tmp_path, monkeypatch, capsys):
    src = _make_docx(tmp_path / "disc.docx")
    out_dir = tmp_path / "pdfs"

    def render(_input, output, options=None):
        with open(output, "wb") as stream:
            stream.write(FAKE_PDF)
        return str(output)

    monkeypatch.setattr(C, "_convert_weasyprint", render)
    monkeypatch.setattr(cli.glob, "glob", lambda pattern: [str(src)])
    rc = cli.main(["--output-dir", str(out_dir), "--engine", "weasyprint"])

    assert rc == 0
    assert "auto-discovered" in capsys.readouterr().err


def test_cli_batch_nonexistent_file_errors(tmp_path):
    with pytest.raises(SystemExit) as exc:
        cli.main(["nonexistent.docx", "--output-dir", str(tmp_path / "pdfs")])
    assert exc.value.code == 2
