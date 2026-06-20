"""Tests de las mejoras de fidelidad: resaltado, mayúsculas/versalitas,
glifos de viñeta, estilo por defecto, metadatos del PDF y validación de entrada.
"""
import pytest

from docx2pdf_py import convert
from docx2pdf_py.converter import Converter
from tests.conftest import NS, document


# -- resaltado (w:highlight) ------------------------------------------------
def test_highlight_named_color(make_docx):
    body = ('<w:p><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr>'
            "<w:t>marcado</w:t></w:r></w:p>")
    with Converter(make_docx(document(body))) as conv:
        html = conv.build_html()
    assert "background-color:#FFFF00" in html


def test_highlight_none_ignored(make_docx):
    body = ('<w:p><w:r><w:rPr><w:highlight w:val="none"/></w:rPr>'
            "<w:t>x</w:t></w:r></w:p>")
    with Converter(make_docx(document(body))) as conv:
        html = conv.build_html()
    assert "background-color" not in html


# -- mayúsculas / versalitas ------------------------------------------------
def test_caps(make_docx):
    body = ('<w:p><w:r><w:rPr><w:caps/></w:rPr><w:t>abc</w:t></w:r></w:p>')
    with Converter(make_docx(document(body))) as conv:
        html = conv.build_html()
    assert "text-transform:uppercase" in html


def test_smallcaps(make_docx):
    body = ('<w:p><w:r><w:rPr><w:smallCaps/></w:rPr><w:t>abc</w:t></w:r></w:p>')
    with Converter(make_docx(document(body))) as conv:
        html = conv.build_html()
    assert "font-variant:small-caps" in html


def test_caps_overrides_smallcaps(make_docx):
    body = ('<w:p><w:r><w:rPr><w:caps/><w:smallCaps/></w:rPr>'
            "<w:t>abc</w:t></w:r></w:p>")
    with Converter(make_docx(document(body))) as conv:
        html = conv.build_html()
    assert "text-transform:uppercase" in html
    assert "small-caps" not in html


# -- glifos de viñeta -------------------------------------------------------
def _bullet_numbering(glyph):
    return (
        f'<w:numbering {NS}>'
        f'<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">'
        f'<w:numFmt w:val="bullet"/><w:lvlText w:val="{glyph}"/></w:lvl>'
        f'</w:abstractNum>'
        f'<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
        f'</w:numbering>'
    )


def _bullet_para():
    return ('<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/>'
            '<w:numId w:val="1"/></w:numPr></w:pPr>'
            "<w:r><w:t>item</w:t></w:r></w:p>")


def test_bullet_glyph_mapped(make_docx):
    # "" (Wingdings) -> "•"
    parts = {"word/numbering.xml": _bullet_numbering("")}
    with Converter(make_docx(document(_bullet_para()), parts=parts)) as conv:
        html = conv.build_html()
    assert "•" in html


def test_bullet_glyph_unknown_falls_back(make_docx):
    parts = {"word/numbering.xml": _bullet_numbering("?")}
    with Converter(make_docx(document(_bullet_para()), parts=parts)) as conv:
        html = conv.build_html()
    assert "•" in html


# -- estilo de párrafo por defecto (w:default="1") --------------------------
def test_default_paragraph_style_applied(make_docx):
    styles = (
        f"<w:styles {NS}>"
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:rPr><w:color w:val="112233"/></w:rPr></w:style>'
        "</w:styles>"
    )
    # párrafo SIN pStyle explícito -> debe heredar el estilo por defecto
    body = "<w:p><w:r><w:t>hola</w:t></w:r></w:p>"
    with Converter(make_docx(document(body), styles_xml=styles)) as conv:
        html = conv.build_html()
    assert "color:#112233" in html


def test_explicit_style_overrides_default(make_docx):
    styles = (
        f"<w:styles {NS}>"
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:rPr><w:color w:val="111111"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Other">'
        '<w:rPr><w:color w:val="222222"/></w:rPr></w:style>'
        "</w:styles>"
    )
    body = ('<w:p><w:pPr><w:pStyle w:val="Other"/></w:pPr>'
            "<w:r><w:t>hola</w:t></w:r></w:p>")
    with Converter(make_docx(document(body), styles_xml=styles)) as conv:
        html = conv.build_html()
    assert "color:#222222" in html


# -- metadatos del documento -> PDF -----------------------------------------
def _core(title="Mi Título", creator="Ada", keywords="a, b"):
    dc = "http://purl.org/dc/elements/1.1/"
    cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    return (
        f'<cp:coreProperties xmlns:cp="{cp}" xmlns:dc="{dc}">'
        f"<dc:title>{title}</dc:title>"
        f"<dc:creator>{creator}</dc:creator>"
        f"<cp:keywords>{keywords}</cp:keywords>"
        "</cp:coreProperties>"
    )


def test_metadata_emitted(make_docx):
    body = "<w:p><w:r><w:t>x</w:t></w:r></w:p>"
    parts = {"docProps/core.xml": _core()}
    with Converter(make_docx(document(body), parts=parts)) as conv:
        html = conv.build_html()
    assert "<title>Mi Título</title>" in html
    assert "name='author' content='Ada'" in html
    assert "name='keywords' content='a, b'" in html


def test_metadata_absent_is_safe(make_docx):
    body = "<w:p><w:r><w:t>x</w:t></w:r></w:p>"
    with Converter(make_docx(document(body))) as conv:
        html = conv.build_html()
    assert "<title>" not in html


# -- validación de entrada en convert() -------------------------------------
def test_convert_missing_file():
    with pytest.raises(FileNotFoundError):
        convert("/no/existe.docx", "out.pdf")


def test_convert_not_a_zip(tmp_path):
    bogus = tmp_path / "fake.docx"
    bogus.write_text("esto no es un zip")
    with pytest.raises(ValueError):
        convert(str(bogus), str(tmp_path / "out.pdf"))
