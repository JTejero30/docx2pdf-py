#!/usr/bin/env python3
"""Smoke test de extremo a extremo: crea un .docx mínimo *completo* (paquete
OOXML que LibreOffice acepta), lo convierte a PDF y verifica el resultado.

Pensado para CI con LibreOffice instalado. Si no hay 'soffice' disponible, se
salta (no falla), para poder ejecutarlo también en local sin LibreOffice.
"""
import sys
import tempfile
import zipfile
from pathlib import Path

from docx2pdf_py import convert, default_engine
from docx2pdf_py.engines import find_libreoffice

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PKG = "http://schemas.openxmlformats.org/package/2006"
OFFICE = "http://schemas.openxmlformats.org/officeDocument/2006"
DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

CONTENT_TYPES = DECL + (
    f'<Types xmlns="{PKG}/content-types">'
    '<Default Extension="rels" '
    'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument'
    '.wordprocessingml.document.main+xml"/>'
    "</Types>"
)

ROOT_RELS = DECL + (
    f'<Relationships xmlns="{PKG}/relationships">'
    f'<Relationship Id="rId1" Type="{OFFICE}/relationships/officeDocument" '
    'Target="word/document.xml"/>'
    "</Relationships>"
)

DOCUMENT_RELS = DECL + f'<Relationships xmlns="{PKG}/relationships"></Relationships>'

DOCUMENT = DECL + (
    f'<w:document xmlns:w="{W}"><w:body>'
    "<w:p><w:r><w:t>Hola, mundo desde docx2pdf-py.</w:t></w:r></w:p>"
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>'
    "</w:body></w:document>"
)


def make_full_docx(path: Path) -> None:
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOCUMENT_RELS)
        z.writestr("word/document.xml", DOCUMENT)


def main() -> int:
    if not find_libreoffice():
        print("LibreOffice no disponible; se omite el smoke test e2e.")
        return 0
    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "sample.docx"
        out = Path(tmp) / "sample.pdf"
        make_full_docx(src)
        print(f"Motor por defecto: {default_engine()}")
        convert(str(src), str(out), engine="libreoffice")
        data = out.read_bytes()
        assert data.startswith(b"%PDF"), "la salida no es un PDF válido"
        assert len(data) > 500, "el PDF parece vacío"
        print(f"OK: PDF generado ({len(data)} bytes).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
