"""Tests using complete OOXML .docx packages (all required boilerplate parts).

These use the make_full_docx fixture, which produces files that real layout
engines (LibreOffice, Word) would accept — unlike the minimal fixtures used in
unit tests, which omit [Content_Types].xml and relationship roots.
"""
from docx2pdf_py.converter import Converter
from tests.conftest import NS

# ---------------------------------------------------------------------------
# Full-package round-trip: build_html on a complete OOXML package
# ---------------------------------------------------------------------------

def test_full_package_basic_text(make_full_docx):
    """A complete OOXML package with plain text converts without error."""
    path = make_full_docx(
        "<w:p><w:r><w:t>Hello from a full OOXML package.</w:t></w:r></w:p>"
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>'
    )
    with Converter(path) as conv:
        html = conv.build_html()
    assert "Hello from a full OOXML package." in html
    assert html.startswith("<!DOCTYPE html>")


def test_full_package_multipage_content(make_full_docx):
    """Many paragraphs with a page-break hint produce break-before:page."""
    paragraphs = "<w:p><w:r><w:t>Page one</w:t></w:r></w:p>"
    paragraphs += (
        "<w:p><w:r><w:lastRenderedPageBreak/>"
        "<w:t>Page two</w:t></w:r></w:p>"
    )
    path = make_full_docx(paragraphs)
    with Converter(path) as conv:
        html = conv.build_html()
    assert "Page one" in html
    assert "Page two" in html
    assert "break-before:page" in html


def test_full_package_table_and_text(make_full_docx):
    """A document mixing paragraphs and a table converts correctly."""
    body = (
        "<w:p><w:r><w:t>Before table</w:t></w:r></w:p>"
        "<w:tbl>"
        '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>'
        "<w:tr>"
        "<w:tc><w:p><w:r><w:t>Cell A</w:t></w:r></w:p></w:tc>"
        "<w:tc><w:p><w:r><w:t>Cell B</w:t></w:r></w:p></w:tc>"
        "</w:tr>"
        "</w:tbl>"
        "<w:p><w:r><w:t>After table</w:t></w:r></w:p>"
    )
    path = make_full_docx(body)
    with Converter(path) as conv:
        html = conv.build_html()
    assert "Before table" in html
    assert "Cell A" in html and "Cell B" in html
    assert "After table" in html
    assert "<table" in html


def test_full_package_mixed_formatting(make_full_docx):
    """A document with bold, italic, colour, and font size converts correctly."""
    body = (
        "<w:p>"
        "<w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r>"
        "<w:r><w:rPr><w:i/></w:rPr><w:t> Italic</w:t></w:r>"
        '<w:r><w:rPr><w:color w:val="CC0000"/><w:sz w:val="28"/></w:rPr>'
        "<w:t> Red14pt</w:t></w:r>"
        "</w:p>"
    )
    path = make_full_docx(body)
    with Converter(path) as conv:
        html = conv.build_html()
    assert "font-weight:bold" in html
    assert "font-style:italic" in html
    assert "color:#CC0000" in html
    assert "font-size:14.0pt" in html


def test_full_package_numbered_list(make_full_docx):
    """A numbered list in a full OOXML package renders with counters."""
    numbering = (
        f'<w:numbering {NS}>'
        '<w:abstractNum w:abstractNumId="0">'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/>'
        '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>'
        '</w:abstractNum>'
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
        '</w:numbering>'
    )
    body = (
        '<w:p><w:pPr><w:numPr>'
        '<w:ilvl w:val="0"/><w:numId w:val="1"/>'
        '</w:numPr></w:pPr><w:r><w:t>First item</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:numPr>'
        '<w:ilvl w:val="0"/><w:numId w:val="1"/>'
        '</w:numPr></w:pPr><w:r><w:t>Second item</w:t></w:r></w:p>'
    )
    path = make_full_docx(body, extra_parts={"word/numbering.xml": numbering})
    with Converter(path) as conv:
        html = conv.build_html()
    assert "First item" in html
    assert "Second item" in html
    assert "1." in html and "2." in html


def test_full_package_nested_table(make_full_docx):
    """A table nested inside another table renders both correctly."""
    inner = (
        "<w:tbl>"
        '<w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>'
        "<w:tr><w:tc><w:p><w:r><w:t>inner</w:t></w:r></w:p></w:tc></w:tr>"
        "</w:tbl>"
    )
    body = (
        "<w:tbl>"
        '<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>'
        "<w:tr><w:tc>"
        "<w:p><w:r><w:t>outer</w:t></w:r></w:p>"
        + inner
        + "</w:tc></w:tr>"
        "</w:tbl>"
    )
    path = make_full_docx(body)
    with Converter(path) as conv:
        html = conv.build_html()
    assert html.count("<table") == 2
    assert "outer" in html and "inner" in html


def test_full_package_colspan_and_rowspan(make_full_docx):
    """A table with both colspan and rowspan renders the correct HTML attributes."""
    body = (
        "<w:tbl>"
        '<w:tblGrid>'
        '<w:gridCol w:w="2000"/><w:gridCol w:w="2000"/>'
        "</w:tblGrid>"
        # Row 1: cell spanning 2 columns
        "<w:tr>"
        '<w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>'
        "<w:p><w:r><w:t>wide</w:t></w:r></w:p></w:tc>"
        "</w:tr>"
        # Row 2: left cell spanning 2 rows vertically (restart), right cell normal
        "<w:tr>"
        '<w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>'
        "<w:p><w:r><w:t>tall</w:t></w:r></w:p></w:tc>"
        "<w:tc><w:p><w:r><w:t>right1</w:t></w:r></w:p></w:tc>"
        "</w:tr>"
        # Row 3: left cell is vMerge continue
        "<w:tr>"
        "<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>"
        "<w:tc><w:p><w:r><w:t>right2</w:t></w:r></w:p></w:tc>"
        "</w:tr>"
        "</w:tbl>"
    )
    path = make_full_docx(body)
    with Converter(path) as conv:
        html = conv.build_html()
    assert 'colspan="2"' in html
    assert 'rowspan="2"' in html
    assert "wide" in html and "tall" in html


def test_full_package_text_alignment(make_full_docx):
    """Paragraph alignment (center, right, justify) maps to CSS text-align."""
    body = (
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
        "<w:r><w:t>centered</w:t></w:r></w:p>"
        '<w:p><w:pPr><w:jc w:val="right"/></w:pPr>'
        "<w:r><w:t>right-aligned</w:t></w:r></w:p>"
        '<w:p><w:pPr><w:jc w:val="both"/></w:pPr>'
        "<w:r><w:t>justified</w:t></w:r></w:p>"
    )
    path = make_full_docx(body)
    with Converter(path) as conv:
        html = conv.build_html()
    assert "text-align:center" in html
    assert "text-align:right" in html
    assert "text-align:justify" in html


def test_full_package_page_size_landscape(make_full_docx):
    """A landscape sectPr emits correct page dimensions in CSS."""
    body = (
        "<w:p><w:r><w:t>landscape</w:t></w:r></w:p>"
        '<w:sectPr>'
        '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
        '<w:pgMar w:top="720" w:bottom="720" w:left="720" w:right="720"/>'
        "</w:sectPr>"
    )
    path = make_full_docx(body)
    with Converter(path) as conv:
        html = conv.build_html()
    assert "size: 29.70cm 21.00cm" in html


def test_full_package_metadata(make_full_docx):
    """Document metadata from docProps/core.xml appears in the HTML <head>."""
    dc = "http://purl.org/dc/elements/1.1/"
    cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    core = (
        f'<cp:coreProperties xmlns:cp="{cp}" xmlns:dc="{dc}">'
        "<dc:title>Test Document</dc:title>"
        "<dc:creator>Test Author</dc:creator>"
        "</cp:coreProperties>"
    )
    path = make_full_docx(
        "<w:p><w:r><w:t>body</w:t></w:r></w:p>",
        extra_parts={"docProps/core.xml": core},
    )
    with Converter(path) as conv:
        html = conv.build_html()
    assert "<title>Test Document</title>" in html
    assert "name='author' content='Test Author'" in html
