"""Tests de las funciones mayores: numeración, ajuste de imágenes y cabeceras
de primera página / pares-impares."""
from docx2pdf_py.converter import Converter, _to_letter, _to_roman
from tests.conftest import NS, document


# ---------------------------------------------------------------------------
# Numeración de listas
# ---------------------------------------------------------------------------
def numbering(fmt="decimal", text="%1.", start="1"):
    return (
        f'<w:numbering {NS}>'
        f'<w:abstractNum w:abstractNumId="0">'
        f'<w:lvl w:ilvl="0"><w:start w:val="{start}"/>'
        f'<w:numFmt w:val="{fmt}"/><w:lvlText w:val="{text}"/></w:lvl>'
        f'<w:lvl w:ilvl="1"><w:start w:val="1"/>'
        f'<w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/></w:lvl>'
        f'</w:abstractNum>'
        f'<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
        f'</w:numbering>'
    )


def list_para(text, ilvl=0, num_id=1):
    return (
        f'<w:p><w:pPr><w:numPr>'
        f'<w:ilvl w:val="{ilvl}"/><w:numId w:val="{num_id}"/>'
        f'</w:numPr></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>'
    )


def test_helpers_letter_roman():
    assert _to_letter(1) == "A"
    assert _to_letter(27) == "AA"
    assert _to_roman(4) == "IV"
    assert _to_roman(2026) == "MMXXVI"


def test_decimal_numbering(make_docx):
    body = list_para("uno") + list_para("dos") + list_para("tres")
    parts = {"word/numbering.xml": numbering()}
    with Converter(make_docx(document(body), parts=parts)) as conv:
        html = conv.build_html()
    # el marcador precede al <span> con el texto del run
    assert "1. <span" in html and "uno" in html
    assert "2. <span" in html and "dos" in html
    assert "3. <span" in html and "tres" in html
    assert "– <span" not in html  # ya no cae a viñeta


def test_nested_numbering_counter_reset(make_docx):
    body = (
        list_para("uno")                 # 1.
        + list_para("hijo a", ilvl=1)    # a)
        + list_para("hijo b", ilvl=1)    # b)
        + list_para("dos")               # 2.
        + list_para("hijo a2", ilvl=1)   # a)  (se reinicia)
    )
    parts = {"word/numbering.xml": numbering()}
    with Converter(make_docx(document(body), parts=parts)) as conv:
        html = conv.build_html()
    assert "1. <span" in html and "uno" in html
    assert "a) <span" in html and "hijo a" in html
    assert "b) <span" in html and "hijo b" in html
    assert "2. <span" in html and "dos" in html
    assert "hijo a2" in html
    # el nivel hijo se reinicia: hay dos marcadores "a)" y ninguno "c)"
    assert html.count("a) <span") == 2
    assert "c) <span" not in html


def test_lower_letter_format(make_docx):
    parts = {"word/numbering.xml": numbering(fmt="lowerLetter", text="%1)")}
    body = list_para("alfa") + list_para("beta")
    with Converter(make_docx(document(body), parts=parts)) as conv:
        html = conv.build_html()
    assert "a) <span" in html and "alfa" in html
    assert "b) <span" in html and "beta" in html


def test_bullet_uses_level_glyph(make_docx):
    parts = {"word/numbering.xml": numbering(fmt="bullet", text="•")}
    with Converter(make_docx(document(list_para("item")), parts=parts)) as conv:
        html = conv.build_html()
    assert "• <span" in html and "item" in html


def test_list_without_numbering_part(make_docx):
    # sin numbering.xml, una lista sigue mostrando viñeta (•)
    with Converter(make_docx(document(list_para("x")))) as conv:
        html = conv.build_html()
    assert "• <span" in html and ">x<" in html


# ---------------------------------------------------------------------------
# Imágenes flotantes con ajuste de texto
# ---------------------------------------------------------------------------
IMG_RELS = (
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://x/image" Target="media/i.png"/>'
    "</Relationships>"
)


def anchored(wrap="wrapSquare", align="left"):
    wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    return (
        f'<w:p><w:r><w:drawing xmlns:wp="{wp}" xmlns:a="{a}">'
        f'<wp:anchor>'
        f'<wp:extent cx="914400" cy="914400"/>'
        f'<wp:positionH relativeFrom="column"><wp:align>{align}</wp:align></wp:positionH>'
        f'<wp:{wrap}/>'
        f'<a:graphic><a:graphicData><a:blip r:embed="rId1"/></a:graphicData></a:graphic>'
        f'</wp:anchor></w:drawing></w:r>'
        f'<w:r><w:t>texto que rodea</w:t></w:r></w:p>'
    )


def test_floating_image_square_wrap(make_docx):
    parts = {"word/_rels/document.xml.rels": IMG_RELS,
             "word/media/i.png": b"\x89PNG\r\n\x1a\n"}
    with Converter(make_docx(document(anchored("wrapSquare", "left")), parts=parts)) as conv:
        html = conv.build_html()
    assert "float:left" in html
    # la imagen va dentro del mismo <p> que el texto (para que lo rodee)
    assert html.index("float:left") < html.index("texto que rodea")


def test_floating_image_right_align(make_docx):
    parts = {"word/_rels/document.xml.rels": IMG_RELS,
             "word/media/i.png": b"\x89PNG\r\n\x1a\n"}
    with Converter(make_docx(document(anchored("wrapTight", "right")), parts=parts)) as conv:
        html = conv.build_html()
    assert "float:right" in html


def test_floating_image_top_bottom_is_block(make_docx):
    parts = {"word/_rels/document.xml.rels": IMG_RELS,
             "word/media/i.png": b"\x89PNG\r\n\x1a\n"}
    with Converter(make_docx(document(anchored("wrapTopAndBottom")), parts=parts)) as conv:
        html = conv.build_html()
    assert "float:" not in html
    assert "display:block" in html


# ---------------------------------------------------------------------------
# Cabeceras de primera página / pares-impares
# ---------------------------------------------------------------------------
def hf_rels():
    return (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rH1" Type="http://x/header" Target="header1.xml"/>'
        '<Relationship Id="rH2" Type="http://x/header" Target="header2.xml"/>'
        '<Relationship Id="rH3" Type="http://x/header" Target="header3.xml"/>'
        "</Relationships>"
    )


def header_part(text):
    return f'<w:hdr {NS}><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:hdr>'


def test_title_page_header(make_docx):
    body = (
        "<w:p><w:r><w:t>cuerpo</w:t></w:r></w:p>"
        '<w:sectPr><w:titlePg/>'
        '<w:headerReference w:type="default" r:id="rH1"/>'
        '<w:headerReference w:type="first" r:id="rH2"/>'
        "</w:sectPr>"
    )
    parts = {
        "word/_rels/document.xml.rels": hf_rels(),
        "word/header1.xml": header_part("DEFECTO"),
        "word/header2.xml": header_part("PRIMERA"),
    }
    with Converter(make_docx(document(body), parts=parts)) as conv:
        html = conv.build_html()
    assert "@page :first" in html
    assert "element(hdr_first)" in html
    assert "PRIMERA" in html
    assert "DEFECTO" in html


def test_title_page_blank_first_header(make_docx):
    # titlePg sin header "first" -> la primera página no muestra cabecera
    body = (
        "<w:p><w:r><w:t>cuerpo</w:t></w:r></w:p>"
        '<w:sectPr><w:titlePg/>'
        '<w:headerReference w:type="default" r:id="rH1"/>'
        "</w:sectPr>"
    )
    parts = {
        "word/_rels/document.xml.rels": hf_rels(),
        "word/header1.xml": header_part("DEFECTO"),
    }
    with Converter(make_docx(document(body), parts=parts)) as conv:
        html = conv.build_html()
    assert "@page :first" in html
    assert "content: none" in html


def test_even_odd_headers(make_docx):
    body = (
        "<w:p><w:r><w:t>cuerpo</w:t></w:r></w:p>"
        '<w:sectPr>'
        '<w:headerReference w:type="default" r:id="rH1"/>'
        '<w:headerReference w:type="even" r:id="rH3"/>'
        "</w:sectPr>"
    )
    parts = {
        "word/_rels/document.xml.rels": hf_rels(),
        "word/header1.xml": header_part("IMPAR"),
        "word/header3.xml": header_part("PAR"),
        "word/settings.xml": f'<w:settings {NS}><w:evenAndOddHeaders/></w:settings>',
    }
    with Converter(make_docx(document(body), parts=parts)) as conv:
        html = conv.build_html()
    assert "@page :left" in html
    assert "element(hdr_even)" in html
    assert "PAR" in html
