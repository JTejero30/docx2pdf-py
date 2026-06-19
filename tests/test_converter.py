"""Tests de docx2pdf_py centrados en build_html (no requieren WeasyPrint)."""
import zipfile

import pytest

from docx2pdf_py import converter as C
from docx2pdf_py.converter import Converter, font_stack
from tests.conftest import document


# -- helpers de formato (puras) ---------------------------------------------
def test_font_stack_metric_substitute():
    assert font_stack("Calibri").startswith("Carlito")
    assert "Gelasio" in font_stack("Georgia")


def test_font_stack_generic_family():
    # una fuente serif conocida no debe caer en sans-serif
    assert font_stack("Times New Roman").endswith("serif")
    assert font_stack("Times New Roman").count("sans-serif") == 0
    assert font_stack("Courier New").endswith("monospace")
    # desconocida -> sans-serif
    assert font_stack("FuenteRara").endswith("sans-serif")


# -- párrafos / runs --------------------------------------------------------
def test_basic_paragraph_text(make_docx):
    path = make_docx(document("<w:p><w:r><w:t>Hola mundo</w:t></w:r></w:p>"))
    with Converter(path) as conv:
        html = conv.build_html()
    assert "Hola mundo" in html
    assert html.startswith("<!DOCTYPE html>")


def test_run_formatting(make_docx):
    body = (
        "<w:p><w:r><w:rPr><w:b/><w:i/>"
        '<w:color w:val="FF0000"/></w:rPr>'
        "<w:t>fuerte</w:t></w:r></w:p>"
    )
    with Converter(make_docx(document(body))) as conv:
        html = conv.build_html()
    assert "font-weight:bold" in html
    assert "font-style:italic" in html
    assert "color:#FF0000" in html


def test_field_value_is_skipped(make_docx):
    # un campo PAGE con valor cacheado "7" no debe imprimirse
    body = (
        "<w:p>"
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        "<w:r><w:instrText> PAGE </w:instrText></w:r>"
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        "<w:r><w:t>7</w:t></w:r>"
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        "</w:p>"
    )
    with Converter(make_docx(document(body))) as conv:
        html = conv.build_html()
    assert ">7<" not in html


def test_page_break(make_docx):
    body = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
    with Converter(make_docx(document(body))) as conv:
        html = conv.build_html()
    assert "break-after:page" in html


# -- tablas -----------------------------------------------------------------
def test_table_gridspan(make_docx):
    body = (
        "<w:tbl>"
        '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>'
        "<w:tr>"
        '<w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>'
        "<w:p><w:r><w:t>cabecera</w:t></w:r></w:p></w:tc>"
        "</w:tr>"
        "</w:tbl>"
    )
    with Converter(make_docx(document(body))) as conv:
        html = conv.build_html()
    assert 'colspan="2"' in html
    assert "<table" in html


def test_table_vmerge_rowspan(make_docx):
    body = (
        "<w:tbl>"
        '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>'
        "<w:tr>"
        '<w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>'
        "<w:p><w:r><w:t>fusion</w:t></w:r></w:p></w:tc>"
        "<w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>"
        "</w:tr>"
        "<w:tr>"
        "<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>"
        "<w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc>"
        "</w:tr>"
        "</w:tbl>"
    )
    with Converter(make_docx(document(body))) as conv:
        html = conv.build_html()
    assert 'rowspan="2"' in html
    # la celda "continue" no debe emitir un segundo "fusion"
    assert html.count("fusion") == 1


# -- hyperlinks -------------------------------------------------------------
def test_hyperlink_href_resolved(make_docx):
    rels = (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://x/hyperlink" '
        'Target="https://example.com" TargetMode="External"/>'
        "</Relationships>"
    )
    body = (
        '<w:p><w:hyperlink r:id="rId1">'
        "<w:r><w:t>enlace</w:t></w:r></w:hyperlink></w:p>"
    )
    parts = {"word/_rels/document.xml.rels": rels}
    with Converter(make_docx(document(body), parts=parts)) as conv:
        html = conv.build_html()
    assert 'href="https://example.com"' in html


# -- tamaño de página -------------------------------------------------------
def test_page_size_from_sectpr_landscape(make_docx):
    # A4 apaisado: 16838 x 11906 twips
    body = (
        '<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
        '<w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/>'
        "</w:sectPr>"
    )
    with Converter(make_docx(document(body))) as conv:
        html = conv.build_html()
    assert "size: 29.70cm 21.00cm" in html  # ancho > alto -> apaisado


def test_page_size_defaults_to_a4_portrait(make_docx):
    with Converter(make_docx(document("<w:p/>"))) as conv:
        html = conv.build_html()
    assert "size: 21.00cm 29.70cm" in html


# -- recursos / seguridad ---------------------------------------------------
def test_context_manager_closes_zip(make_docx):
    conv = Converter(make_docx(document("<w:p/>")))
    conv.close()
    assert conv.z is None


def test_xxe_entity_not_resolved(make_docx):
    # una entidad externa no debe resolverse (parser endurecido)
    doc = (
        '<?xml version="1.0"?>'
        '<!DOCTYPE w:document [<!ENTITY xxe "INJECTED">]>'
        + document("<w:p><w:r><w:t>&xxe;</w:t></w:r></w:p>")
    )
    # con resolve_entities=False, lxml deja la entidad sin expandir
    with Converter(make_docx(doc)) as conv:
        html = conv.build_html()
    assert "INJECTED" not in html


def test_zip_bomb_member_limit(make_docx, monkeypatch):
    monkeypatch.setattr(C, "MAX_MEMBER_BYTES", 10)
    path = make_docx(document("<w:p><w:r><w:t>texto bastante largo</w:t></w:r></w:p>"))
    with pytest.raises(ValueError):
        Converter(path)
