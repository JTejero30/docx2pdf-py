"""Conversion backends using real layout engines for .docx -> PDF.

A .docx does not store fixed pages: they are computed by the layout engine at
render time. The native flow (lxml + WeasyPrint) can only *approximate* Word's
pagination. When a real engine is available the PDF has the **same content per
page** as the original document:

- **Microsoft Word** (Windows via COM, macOS via AppleScript): identical
  pagination to Word.
- **LibreOffice** (headless, cross-platform): faithful pagination as LibreOffice
  renders the document (a different engine to Word, very similar but not
  guaranteed identical).

If neither is available, the caller falls back to the lxml + WeasyPrint flow.
"""
import logging
import os
import platform
import shutil
import subprocess
import sys
import tempfile
from typing import Optional

from .exceptions import ConversionError, ConversionTimeoutError, EngineUnavailableError
from .output import Pathish, publish_pdf
from .processes import run_process

logger = logging.getLogger(__name__)

# Ejecutables de LibreOffice a buscar en el PATH, y rutas habituales de
# instalación en macOS/Windows donde el binario no suele estar en el PATH.
_SOFFICE_NAMES = ("soffice", "libreoffice")
_SOFFICE_PATHS = (
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
)


# -- LibreOffice ------------------------------------------------------------
def find_libreoffice() -> Optional[str]:
    """Path to the LibreOffice executable (``soffice``), or ``None`` if absent.

    Can be overridden with the ``SOFFICE_BIN`` environment variable.
    """
    override = os.environ.get("SOFFICE_BIN")
    if override and os.path.exists(override):
        return override
    for name in _SOFFICE_NAMES:
        path = shutil.which(name)
        if path:
            return path
    for path in _SOFFICE_PATHS:
        if os.path.exists(path):
            return path
    return None


def convert_libreoffice(
    in_path: Pathish, out_path: Pathish, soffice: Optional[str] = None, timeout: int = 120
) -> str:
    """Convert ``in_path`` -> ``out_path`` using LibreOffice headless."""
    soffice = soffice or find_libreoffice()
    if not soffice:
        raise EngineUnavailableError("LibreOffice (soffice) is not available")
    in_path = os.path.abspath(in_path)
    out_path = os.path.abspath(out_path)
    with tempfile.TemporaryDirectory() as tmp:
        # Isolated user profile: allows concurrent runs and avoids colliding
        # with an existing LibreOffice instance the user may have open.
        profile = os.path.join(tmp, "profile")
        cmd = [
            soffice, "--headless", "--norestore",
            "-env:UserInstallation=file://" + profile,
            "--convert-to", "pdf", "--outdir", tmp, in_path,
        ]
        try:
            proc = run_process(cmd, timeout=timeout or None)
        except subprocess.TimeoutExpired as exc:
            raise ConversionTimeoutError(
                f"LibreOffice timed out after {timeout}s"
            ) from exc
        produced = os.path.join(
            tmp, os.path.splitext(os.path.basename(in_path))[0] + ".pdf"
        )
        if proc.returncode != 0 or not os.path.exists(produced):
            detail = (proc.stderr or proc.stdout or b"").decode(errors="replace").strip()
            raise ConversionError("LibreOffice failed to convert the document" +
                                  (f": {detail}" if detail else ""))
        return publish_pdf(produced, out_path)


# -- Microsoft Word ---------------------------------------------------------
def word_available() -> bool:
    """Return True if Microsoft Word can be automated on this system."""
    system = platform.system()
    if system == "Windows":
        try:
            import winreg

            open_key = getattr(winreg, "OpenKey", None)
            classes_root = getattr(winreg, "HKEY_CLASSES_ROOT", None)
            if open_key is None or classes_root is None:
                return False
            with open_key(classes_root, r"Word.Application\CLSID"):
                installed = True
        except (ImportError, FileNotFoundError, OSError):
            installed = False
        if not installed:
            return False
        for mod in ("win32com.client", "comtypes.client"):
            try:
                __import__(mod)
                return True
            except Exception:
                continue
        return False
    if system == "Darwin":
        return os.path.exists("/Applications/Microsoft Word.app")
    return False


def convert_word(in_path: Pathish, out_path: Pathish, timeout: int = 120) -> str:
    """Convert ``in_path`` -> ``out_path`` by automating Microsoft Word."""
    in_path = os.path.abspath(in_path)
    out_path = os.path.abspath(out_path)
    system = platform.system()
    if system == "Windows":
        with tempfile.TemporaryDirectory() as tmp:
            staged = os.path.join(tmp, "document.pdf")
            command = [
                sys.executable, "-m", "docx2pdf_py._word_worker", in_path, staged
            ]
            try:
                proc = run_process(command, timeout=timeout or None)
            except subprocess.TimeoutExpired as exc:
                raise ConversionTimeoutError(f"Word timed out after {timeout}s") from exc
            if proc.returncode != 0:
                detail = (proc.stderr or proc.stdout).decode(errors="replace").strip()
                raise ConversionError(
                    "Word failed to convert the document"
                    + (f": {detail}" if detail else "")
                )
            return publish_pdf(staged, out_path)
    if system == "Darwin":
        return _convert_word_macos(in_path, out_path, timeout=timeout)
    raise EngineUnavailableError(
        "Microsoft Word automation is only supported on Windows and macOS"
    )


def _convert_word_windows(in_path: str, out_path: str) -> str:
    try:
        import win32com.client as client
        app = client.Dispatch("Word.Application")
    except Exception:
        import comtypes.client as client
        app = client.CreateObject("Word.Application")
    app.Visible = False
    with tempfile.TemporaryDirectory() as tmp:
        staged = os.path.join(tmp, "document.pdf")
        doc = None
        try:
            doc = app.Documents.Open(in_path, ReadOnly=True)
            doc.SaveAs(staged, FileFormat=17)  # 17 = wdFormatPDF
            return publish_pdf(staged, out_path)
        finally:
            if doc is not None:
                try:
                    doc.Close(False)
                except Exception as exc:
                    logger.debug("Word document cleanup failed: %s", exc)
            try:
                app.Quit()
            except Exception as exc:
                logger.debug("Word application cleanup failed: %s", exc)


def _convert_word_macos(in_path: str, out_path: str, timeout: int = 120) -> str:
    # Paths are argv values, avoiding quoting bugs and AppleScript injection.
    script = """on run argv
set inputPath to item 1 of argv
set outputPath to item 2 of argv
tell application "Microsoft Word"
  set theDoc to open file name (POSIX file inputPath as string)
  save as theDoc file name outputPath file format format PDF
  close theDoc saving no
end tell
end run"""
    with tempfile.TemporaryDirectory() as tmp:
        staged = os.path.join(tmp, "document.pdf")
        try:
            proc = run_process(
                ["osascript", "-e", script, "--", in_path, staged],
                timeout=timeout or None,
            )
        except subprocess.TimeoutExpired as exc:
            raise ConversionTimeoutError(f"Word timed out after {timeout}s") from exc
        if proc.returncode != 0 or not os.path.exists(staged):
            detail = proc.stderr.decode(errors="replace").strip()
            raise ConversionError("Word (macOS) failed to convert the document" +
                                  (f": {detail}" if detail else ""))
        return publish_pdf(staged, out_path)


# -- engine selection -------------------------------------------------------
def default_engine() -> str:
    """Return the engine that ``auto`` mode would use on this system.

    Orden de preferencia: LibreOffice -> Word -> WeasyPrint (Python puro).
    """
    if find_libreoffice():
        return "libreoffice"
    if word_available():
        return "word"
    return "weasyprint"
