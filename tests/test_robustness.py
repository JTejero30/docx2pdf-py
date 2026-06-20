"""Edge-case and robustness tests for the OOXML converter.

Covers malformed input, missing optional parts, empty structures, and
boundary conditions that the happy-path tests do not exercise.
"""
import zipfile

import pytest
from lxml import etree

from docx2pdf_py.converter import Converter, _to_letter, _to_roman
from tests.conftest import NS, document

# ---------------------------------------------------------------------------
# Missing / malformed parts
# ---------------------------------------------------------------------------

def test_missing_document_xml_raises(tmp_path):
    """A .docx without word/document.xml must raise ValueError, not KeyError."""
    path = tmp_path / "bad.docx"
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("word/styles.xml", f"<w:styles {NS}/>")
    with pytest.raises(ValueError, match="required OOXML part missing"):
        Converter(str(path))


def test_missing_styles_xml_raises(tmp_path):
    """A .docx without word/styles.xml must raise ValueError, not KeyError."""
    path = tmp_path / "bad.docx"
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("word/document.xml",
                   document("<w:p><w:r><w:t>x</w:t></w:r></w:p>"))
    with pytest.raises(ValueError, match="required OOXML part missing"):
        Converter(str(path))


def test_malformed_xml_raises(tmp_path):
    """Truncated XML must raise (lxml parse error), not hang or crash silently."""
    path = tmp_path / "bad.docx"
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("word/document.xml", "<w:document <unclosed")
        z.writestr("word/styles.xml", f"<w:styles {NS}/>")
    with pytest.raises(etree.XMLSyntaxError):
        Converter(str(path))


# ---------------------------------------------------------------------------
# Optional missing parts (should degrade gracefully)
# ---------------------------------------------------------------------------

def test_missing_numbering_xml_no_crash(make_docx):
    """A list paragraph without numbering.xml falls back to bullet glyph."""
    body = (
        '<w:p><w:pPr><w:numPr>'
        '<w:ilvl w:val="0"/><w:numId w:val="99"/>'
        '</w:numPr></w:pPr><w:r><w:t>item</w:t></w:r></w:p>'
    )
    with Converter(make_docx(document(body))) as conv:
        html = conv.build_html()
    assert "item" in html
    assert "•" in html


def test_missing_rels_file_no_crash(make_docx):
    """A document with no word/_rels/document.xml.rels processes without error."""
    with Converter(make_docx(document("<w:p><w:r><w:t>ok</w:t></w:r></w:p>"))) as conv:
        html = conv.build_html()
    assert "ok" in html


def test_hyperlink_missing_rel_target(make_docx):
    """A hyperlink whose rId has no matching relationship renders without href."""
    body = (
        '<w:p><w:hyperlink r:id="rIdMissing">'
        "<w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>"
    )
    with Converter(make_docx(document(body))) as conv:
        html = conv.build_html()
    assert "link" in html
    assert 'href="' not in html


def test_image_missing_rel_target(make_docx):
    """An image whose relationship target is absent is silently skipped."""
    wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    body = (
        f'<w:p><w:r><w:drawing xmlns:wp="{wp}" xmlns:a="{a}">'
        "<wp:inline>"
        f'<a:graphic><a:graphicData><a:blip xmlns:r="{r_ns}" r:embed="rIdGone"/>'
        "</a:graphicData></a:graphic>"
        "</wp:inline></w:drawing></w:r></w:p>"
    )
    with Converter(make_docx(document(body))) as conv:
        html = conv.build_html()
    assert "<img" not in html


# ---------------------------------------------------------------------------
# Empty structures
# ---------------------------------------------------------------------------

def test_empty_table_no_crash(make_docx):
    """An empty <w:tbl> renders as an empty HTML table without crashing."""
    body = "<w:tbl><w:tblGrid/></w:tbl>"
    with Converter(make_docx(document(body))) as conv:
        html = conv.build_html()
    assert "<table" in html


def test_empty_table_row_no_crash(make_docx):
    """A table with an empty row renders without crashing."""
    body = (
        "<w:tbl>"
        '<w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>'
        "<w:tr></w:tr>"
        "</w:tbl>"
    )
    with Converter(make_docx(document(body))) as conv:
        html = conv.build_html()
    assert "<table" in html


def test_empty_paragraph_renders_space(make_docx):
    """An empty paragraph produces a non-empty <p> (preserves line height)."""
    with Converter(make_docx(document("<w:p/>"))) as conv:
        html = conv.build_html()
    assert "<p " in html


def test_paragraph_only_whitespace(make_docx):
    """A paragraph with only spaces doesn't collapse to nothing."""
    body = "<w:p><w:r><w:t xml:space='preserve'>   </w:t></w:r></w:p>"
    with Converter(make_docx(document(body))) as conv:
        html = conv.build_html()
    assert "<p " in html


# ---------------------------------------------------------------------------
# Style inheritance edge cases
# ---------------------------------------------------------------------------

def test_style_cycle_no_infinite_loop(make_docx):
    """Circular basedOn chains must not cause infinite recursion."""
    styles = (
        f"<w:styles {NS}>"
        '<w:style w:styleId="A"><w:basedOn w:val="B"/>'
        '<w:rPr><w:b/></w:rPr></w:style>'
        '<w:style w:styleId="B"><w:basedOn w:val="A"/>'
        '<w:rPr><w:i/></w:rPr></w:style>'
        "</w:styles>"
    )
    body = '<w:p><w:pPr><w:pStyle w:val="A"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>'
    with Converter(make_docx(document(body), styles_xml=styles)) as conv:
        html = conv.build_html()
    assert "x" in html


def test_style_based_on_nonexistent_parent(make_docx):
    """A style whose basedOn points to a missing style renders without error."""
    styles = (
        f"<w:styles {NS}>"
        '<w:style w:styleId="Child"><w:basedOn w:val="Ghost"/>'
        '<w:rPr><w:color w:val="AABBCC"/></w:rPr></w:style>'
        "</w:styles>"
    )
    body = '<w:p><w:pPr><w:pStyle w:val="Child"/></w:pPr><w:r><w:t>ok</w:t></w:r></w:p>'
    with Converter(make_docx(document(body), styles_xml=styles)) as conv:
        html = conv.build_html()
    assert "color:#AABBCC" in html


# ---------------------------------------------------------------------------
# Text rendering edge cases
# ---------------------------------------------------------------------------

def test_tab_renders_as_spaces(make_docx):
    """<w:tab/> inside a run renders as whitespace, not disappears."""
    body = "<w:p><w:r><w:tab/><w:t>after tab</w:t></w:r></w:p>"
    with Converter(make_docx(document(body))) as conv:
        html = conv.build_html()
    assert "after tab" in html


def test_line_break_renders_br(make_docx):
    """<w:br/> (soft return) renders as <br>."""
    body = "<w:p><w:r><w:t>line1</w:t><w:br/><w:t>line2</w:t></w:r></w:p>"
    with Converter(make_docx(document(body))) as conv:
        html = conv.build_html()
    assert "<br>" in html
    assert "line1" in html and "line2" in html


def test_multiple_runs_in_paragraph(make_docx):
    """Multiple runs with different formatting merge into a single paragraph."""
    body = (
        "<w:p>"
        "<w:r><w:rPr><w:b/></w:rPr><w:t>bold </w:t></w:r>"
        "<w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r>"
        "</w:p>"
    )
    with Converter(make_docx(document(body))) as conv:
        html = conv.build_html()
    assert "bold" in html and "italic" in html
    assert html.count("<p ") == 1


def test_superscript_and_subscript(make_docx):
    """Superscript and subscript vertical alignment renders in CSS."""
    body = (
        "<w:p>"
        '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>sup</w:t></w:r>'
        '<w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>sub</w:t></w:r>'
        "</w:p>"
    )
    with Converter(make_docx(document(body))) as conv:
        html = conv.build_html()
    assert "vertical-align:super" in html
    assert "vertical-align:sub" in html


# ---------------------------------------------------------------------------
# Unit conversion helpers
# ---------------------------------------------------------------------------

def test_to_letter_boundaries():
    assert _to_letter(1) == "A"
    assert _to_letter(26) == "Z"
    assert _to_letter(27) == "AA"
    assert _to_letter(52) == "AZ"
    assert _to_letter(53) == "BA"


def test_to_roman_boundaries():
    assert _to_roman(1) == "I"
    assert _to_roman(4) == "IV"
    assert _to_roman(9) == "IX"
    assert _to_roman(40) == "XL"
    assert _to_roman(400) == "CD"
    assert _to_roman(900) == "CM"
    assert _to_roman(2026) == "MMXXVI"
    assert _to_roman(0) == "0"
    assert _to_roman(-1) == "-1"
