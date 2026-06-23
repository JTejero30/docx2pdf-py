"""Snapshot (golden-file) tests for the OOXML→HTML pipeline.

Run with --update-snapshots to regenerate:
    pytest tests/test_snapshots.py --update-snapshots
"""
from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

SNAPSHOT_DIR = Path(__file__).parent / "snapshots"

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
STYLES = f'<w:styles {NS}></w:styles>'


def _make_docx(tmp_path, body_xml: str, name: str) -> str:
    path = tmp_path / name
    doc = f'<w:document {NS}><w:body>{body_xml}</w:body></w:document>'
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("word/document.xml", doc)
        z.writestr("word/styles.xml", STYLES)
    return str(path)


def _check_snapshot(name: str, actual: str, update: bool) -> None:
    # Las @font-face llevan rutas file:// absolutas (dependen de la máquina);
    # se normalizan para que el snapshot sea portable.
    import re
    actual = re.sub(r"src:url\('file://[^']*'\)", "src:url('BUNDLED')", actual)
    SNAPSHOT_DIR.mkdir(exist_ok=True)
    snap = SNAPSHOT_DIR / name
    if update or not snap.exists():
        snap.write_text(actual, encoding="utf-8")
        return
    expected = snap.read_text(encoding="utf-8")
    assert actual == expected, (
        f"Snapshot mismatch for {name}. "
        "Run pytest --update-snapshots to regenerate."
    )


@pytest.fixture
def update_snapshots(request) -> bool:
    return request.config.getoption("--update-snapshots", default=False)


def test_plain_paragraph_snapshot(tmp_path, update_snapshots):
    from docx2pdf_py.converter import Converter
    body = '<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>'
    path = _make_docx(tmp_path, body, "plain.docx")
    with Converter(path) as c:
        html = c.build_html()
    _check_snapshot("plain_paragraph.html", html, update_snapshots)
    assert "Hello world" in html


def test_bold_italic_snapshot(tmp_path, update_snapshots):
    from docx2pdf_py.converter import Converter
    body = (
        '<w:p><w:r><w:rPr><w:b/><w:i/></w:rPr>'
        '<w:t>Bold and italic</w:t></w:r></w:p>'
    )
    path = _make_docx(tmp_path, body, "bold_italic.docx")
    with Converter(path) as c:
        html = c.build_html()
    _check_snapshot("bold_italic.html", html, update_snapshots)
    assert "font-weight:bold" in html
    assert "font-style:italic" in html


def test_table_snapshot(tmp_path, update_snapshots):
    from docx2pdf_py.converter import Converter
    body = (
        '<w:tbl><w:tr>'
        '<w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>'
        '<w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>'
        '</w:tr></w:tbl>'
    )
    path = _make_docx(tmp_path, body, "table.docx")
    with Converter(path) as c:
        html = c.build_html()
    _check_snapshot("table.html", html, update_snapshots)
    assert "<table" in html
    assert ">A<" in html
