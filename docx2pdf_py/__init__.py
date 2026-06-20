"""docx2pdf_py — conversión fiel de .docx a PDF usando solo librerías de Python.

Lee el OOXML del documento (estilos reales: fuentes, colores, bordes, sombreados,
tablas, imágenes, cabecera/pie) y lo recrea como HTML que WeasyPrint pagina a PDF.

Uso:
    from docx2pdf_py import convert
    convert("entrada.docx", "salida.pdf")
"""
from .converter import Converter, convert
from .engines import default_engine, find_libreoffice, word_available

__version__ = "0.1.0"
__all__ = [
    "convert", "Converter", "__version__",
    "default_engine", "find_libreoffice", "word_available",
]
