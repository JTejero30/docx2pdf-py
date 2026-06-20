"""Utilities for building minimal in-memory .docx files for tests."""
import zipfile

import pytest

NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)

MINIMAL_STYLES = f'<w:styles {NS}></w:styles>'

# Namespace URIs used in full OOXML package boilerplate.
_PKG = "http://schemas.openxmlformats.org/package/2006"
_OFFICE = "http://schemas.openxmlformats.org/officeDocument/2006"
_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

_CONTENT_TYPES = _DECL + (
    f'<Types xmlns="{_PKG}/content-types">'
    '<Default Extension="rels" '
    'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument'
    '.wordprocessingml.document.main+xml"/>'
    "</Types>"
)

_ROOT_RELS = _DECL + (
    f'<Relationships xmlns="{_PKG}/relationships">'
    f'<Relationship Id="rId1" Type="{_OFFICE}/relationships/officeDocument" '
    'Target="word/document.xml"/>'
    "</Relationships>"
)

_DOCUMENT_RELS = _DECL + f'<Relationships xmlns="{_PKG}/relationships"></Relationships>'


def document(body_xml: str) -> str:
    """Wrap body XML in a valid word/document.xml."""
    return f'<w:document {NS}><w:body>{body_xml}</w:body></w:document>'


@pytest.fixture
def make_docx(tmp_path):
    """Return a function that writes a minimal .docx with the given parts.

    The resulting file contains only word/document.xml and word/styles.xml
    (plus any extra ``parts``). It is sufficient for Converter but not for
    real layout engines like LibreOffice.
    """
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


@pytest.fixture
def make_full_docx(tmp_path):
    """Return a function that writes a *complete* OOXML .docx package.

    The file includes all required boilerplate parts so that LibreOffice and
    other real layout engines will accept it. Use this fixture when testing
    integration with the full conversion pipeline.
    """
    counter = {"n": 0}

    def _make(body_xml: str, extra_parts=None):
        counter["n"] += 1
        path = tmp_path / f"full_doc{counter['n']}.docx"
        doc_xml = (
            _DECL
            + f'<w:document xmlns:w="{_W}"><w:body>'
            + body_xml
            + "</w:body></w:document>"
        )
        with zipfile.ZipFile(path, "w") as z:
            z.writestr("[Content_Types].xml", _CONTENT_TYPES)
            z.writestr("_rels/.rels", _ROOT_RELS)
            z.writestr("word/_rels/document.xml.rels", _DOCUMENT_RELS)
            z.writestr("word/document.xml", doc_xml)
            z.writestr("word/styles.xml", MINIMAL_STYLES)
            for name, data in (extra_parts or {}).items():
                z.writestr(name, data)
        return str(path)

    return _make
