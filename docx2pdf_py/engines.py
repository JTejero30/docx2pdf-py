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
import os
import platform
import shutil
import subprocess
import tempfile

# Ejecutables de LibreOffice a buscar en el PATH, y rutas habituales de
# instalación en macOS/Windows donde el binario no suele estar en el PATH.
_SOFFICE_NAMES = ("soffice", "libreoffice")
_SOFFICE_PATHS = (
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
)


# -- LibreOffice ------------------------------------------------------------
def find_libreoffice():
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


def convert_libreoffice(in_path, out_path, soffice=None, timeout=120):
    """Convert ``in_path`` -> ``out_path`` using LibreOffice headless."""
    soffice = soffice or find_libreoffice()
    if not soffice:
        raise RuntimeError("LibreOffice (soffice) is not available")
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
        proc = subprocess.run(cmd, capture_output=True, timeout=timeout)
        produced = os.path.join(
            tmp, os.path.splitext(os.path.basename(in_path))[0] + ".pdf"
        )
        if proc.returncode != 0 or not os.path.exists(produced):
            detail = (proc.stderr or proc.stdout or b"").decode(errors="replace").strip()
            raise RuntimeError("LibreOffice failed to convert the document" +
                               (f": {detail}" if detail else ""))
        os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
        shutil.copyfile(produced, out_path)
    return out_path


# -- Microsoft Word ---------------------------------------------------------
def word_available():
    """Return True if Microsoft Word can be automated on this system."""
    system = platform.system()
    if system == "Windows":
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


def convert_word(in_path, out_path):
    """Convert ``in_path`` -> ``out_path`` by automating Microsoft Word."""
    in_path = os.path.abspath(in_path)
    out_path = os.path.abspath(out_path)
    system = platform.system()
    if system == "Windows":
        return _convert_word_windows(in_path, out_path)
    if system == "Darwin":
        return _convert_word_macos(in_path, out_path)
    raise RuntimeError("Microsoft Word automation is only supported on Windows and macOS")


def _convert_word_windows(in_path, out_path):
    try:
        import win32com.client as client
        app = client.Dispatch("Word.Application")
    except Exception:
        import comtypes.client as client  # type: ignore
        app = client.CreateObject("Word.Application")
    app.Visible = False
    doc = None
    try:
        doc = app.Documents.Open(in_path, ReadOnly=True)
        doc.SaveAs(out_path, FileFormat=17)  # 17 = wdFormatPDF
    finally:
        if doc is not None:
            doc.Close(False)
        app.Quit()
    return out_path


def _convert_word_macos(in_path, out_path):
    # AppleScript: open the document in Word and save it as PDF.
    script = (
        'tell application "Microsoft Word"\n'
        f'  set theDoc to open file name (POSIX file "{in_path}" as string)\n'
        f'  save as theDoc file name "{out_path}" file format format PDF\n'
        '  close theDoc saving no\n'
        'end tell'
    )
    proc = subprocess.run(["osascript", "-e", script], capture_output=True, timeout=120)
    if proc.returncode != 0 or not os.path.exists(out_path):
        detail = proc.stderr.decode(errors="replace").strip()
        raise RuntimeError("Word (macOS) failed to convert the document" +
                           (f": {detail}" if detail else ""))
    return out_path


# -- engine selection -------------------------------------------------------
def default_engine():
    """Return the engine that ``auto`` mode would use on this system."""
    if word_available():
        return "word"
    if find_libreoffice():
        return "libreoffice"
    return "weasyprint"
