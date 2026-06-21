"""Performance benchmarks for the OOXML→HTML conversion pipeline."""
from __future__ import annotations

import zipfile

import pytest

# requires: pip install pytest-benchmark
pytest.importorskip("pytest_benchmark")


NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
STYLES = f'<w:styles {NS}></w:styles>'


def _make_docx(tmp_path, body_xml: str, n: int = 1) -> str:
    path = tmp_path / f"bench{n}.docx"
    doc = f'<w:document {NS}><w:body>{body_xml}</w:body></w:document>'
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("word/document.xml", doc)
        z.writestr("word/styles.xml", STYLES)
    return str(path)


def _para(text: str, bold: bool = False) -> str:
    rpr = "<w:rPr><w:b/></w:rPr>" if bold else ""
    return f"<w:p><w:r>{rpr}<w:t>{text}</w:t></w:r></w:p>"


def _table(rows: int, cols: int) -> str:
    cell = "<w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc>"
    row = "<w:tr>" + cell * cols + "</w:tr>"
    return "<w:tbl>" + row * rows + "</w:tbl>"


@pytest.fixture
def simple_doc(tmp_path):
    body = "".join(_para(f"Paragraph {i}", bold=i % 3 == 0) for i in range(20))
    return _make_docx(tmp_path, body)


@pytest.fixture
def table_doc(tmp_path):
    body = _table(10, 5)
    return _make_docx(tmp_path, body, n=2)


@pytest.fixture
def large_doc(tmp_path):
    body = "".join(_para(f"Line {i} with some content here.", bold=i % 5 == 0) for i in range(100))
    body += _table(20, 4)
    return _make_docx(tmp_path, body, n=3)


def test_build_html_simple(benchmark, simple_doc):
    from docx2pdf_py.converter import Converter
    def run():
        with Converter(simple_doc) as c:
            return c.build_html()
    html = benchmark(run)
    assert "<html>" in html or "<!DOCTYPE" in html


def test_build_html_table(benchmark, table_doc):
    from docx2pdf_py.converter import Converter
    def run():
        with Converter(table_doc) as c:
            return c.build_html()
    html = benchmark(run)
    assert "<table" in html


def test_build_html_large(benchmark, large_doc):
    from docx2pdf_py.converter import Converter
    def run():
        with Converter(large_doc) as c:
            return c.build_html()
    html = benchmark(run)
    assert len(html) > 1000
