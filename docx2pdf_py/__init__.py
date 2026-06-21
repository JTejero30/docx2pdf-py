"""docx2pdf_py — faithful .docx to PDF conversion using pure Python libraries.

Reads the OOXML from the document (real styles: fonts, colours, borders,
shading, tables, images, headers/footers) and recreates it as HTML that
WeasyPrint lays out and paginates into a PDF.

Usage:
    from docx2pdf_py import convert
    convert("input.docx", "output.pdf")
"""
from importlib.metadata import PackageNotFoundError, version

from .api import convert, convert_batch, convert_batch_async, convert_detailed
from .converter import Converter
from .engine_protocol import ConversionEngine
from .engines import default_engine, find_libreoffice, word_available
from .exceptions import (
    ConversionError,
    ConversionTimeoutError,
    Docx2PdfError,
    EngineUnavailableError,
    InvalidDocumentError,
)
from .models import (
    BatchItemResult,
    ConversionAttempt,
    ConversionOptions,
    ConversionResult,
)

try:
    __version__ = version("docx2pdf-py")
except PackageNotFoundError:
    __version__ = "0+unknown"
__all__ = [
    "convert", "convert_detailed", "convert_batch", "convert_batch_async",
    "Converter", "ConversionOptions",
    "ConversionResult", "ConversionAttempt", "BatchItemResult", "ConversionEngine",
    "Docx2PdfError", "InvalidDocumentError", "EngineUnavailableError",
    "ConversionError", "ConversionTimeoutError", "__version__",
    "default_engine", "find_libreoffice", "word_available",
]
