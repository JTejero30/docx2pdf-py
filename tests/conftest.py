"""Utilidades para construir .docx mínimos en memoria para los tests."""
import zipfile

import pytest

NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)

MINIMAL_STYLES = f'<w:styles {NS}></w:styles>'


def document(body_xml: str) -> str:
    """Envuelve el XML de cuerpo en un word/document.xml válido."""
    return f'<w:document {NS}><w:body>{body_xml}</w:body></w:document>'


@pytest.fixture
def make_docx(tmp_path):
    """Devuelve una función que escribe un .docx con las partes dadas."""
    counter = {"n": 0}

    def _make(document_xml: str, styles_xml: str = MINIMAL_STYLES, parts=None):
        counter["n"] += 1
        path = tmp_path / f"doc{counter['n']}.docx"
        with zipfile.ZipFile(path, "w") as z:
            z.writestr("word/document.xml", document_xml)
            z.writestr("word/styles.xml", styles_xml)
            for name, data in (parts or {}).items():
                z.writestr(name, data)
        return str(path)

    return _make
