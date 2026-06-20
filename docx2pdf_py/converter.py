#!/usr/bin/env python3
"""docx -> PDF using pure Python libraries.

Reads OOXML from a .docx (real styles: fonts, colours, borders, shading,
tables, headers/footers) and recreates it as HTML, which WeasyPrint lays out
and paginates into a PDF. Calibri/Georgia are mapped to their metrically
compatible free equivalents Carlito/Gelasio.

Usage:
    from docx2pdf_py import convert
    convert("input.docx", "output.pdf")
"""
import base64
import concurrent.futures
import html as _html
import os
import re
import sys
import zipfile
from typing import Any, Optional

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
DC = "http://purl.org/dc/elements/1.1/"
CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"

# Hardened parser: a .docx is untrusted input. Disable entity resolution
# (XXE / "billion laughs") and network access.
_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False)

# Defensive limits against zip bombs: max decompressed size per member and
# in total (a normal .docx is well below these figures).
MAX_MEMBER_BYTES = 200 * 1024 * 1024
MAX_TOTAL_BYTES = 500 * 1024 * 1024

# Unit-conversion constants.
_EMU_PER_PT: float = 12700.0   # English Metric Units per typographic point
_TWIP_PER_PT: float = 20.0     # twentieths of a point per point
_TWIP_PER_CM: float = 566.929  # twentieths of a point per centimetre

# WeasyPrint rendering timeout in seconds (0 = no timeout).
_WEASYPRINT_TIMEOUT = int(os.environ.get("WEASYPRINT_TIMEOUT", "120"))

BLOCK_IMG_STYLE = "display:block;margin:6pt auto;max-width:100%;"


def _xml(data: bytes):
    return etree.fromstring(data, _PARSER)


def w(tag: str) -> str:
    return f"{{{W}}}{tag}"


def emu_pt(emu) -> float:
    return float(emu) / _EMU_PER_PT


def first(el, tag: str):
    if el is None:
        return None
    return el.find(w(tag))


def attr(el, name: str):
    if el is None:
        return None
    return el.get(w(name))


def on(el):
    """Un elemento booleano OOXML (w:b, w:i, ...) está activo salvo val=0/false."""
    if el is None:
        return None
    v = el.get(w("val"))
    return v not in ("0", "false", "off")


def tw_pt(twips) -> float:
    return float(twips) / _TWIP_PER_PT


def tw_cm(twips) -> float:
    return float(twips) / _TWIP_PER_CM


def esc(s: str) -> str:
    return _html.escape(s, quote=False)


def keep_spaces(s: str) -> str:
    """Conserva espacios múltiples/iniciales (HTML los colapsaría)."""
    s = esc(s)
    s = re.sub(r"  +", lambda m: " " + " " * (len(m.group(0)) - 1), s)
    if s.startswith(" "):
        s = " " + s[1:]
    return s


# -- numeración de listas ----------------------------------------------------
def _to_letter(n: int) -> str:
    """1 -> a, 26 -> z, 27 -> aa (estilo Word lowerLetter/upperLetter)."""
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s or "A"


def _to_roman(n: int) -> str:
    if n <= 0:
        return str(n)
    table = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"),
             (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"),
             (5, "V"), (4, "IV"), (1, "I")]
    out = ""
    for v, sym in table:
        while n >= v:
            out += sym
            n -= v
    return out


def _format_num(n: int, fmt: str) -> str:
    if fmt == "lowerLetter":
        return _to_letter(n).lower()
    if fmt == "upperLetter":
        return _to_letter(n).upper()
    if fmt == "lowerRoman":
        return _to_roman(n).lower()
    if fmt == "upperRoman":
        return _to_roman(n).upper()
    if fmt == "decimalZero":
        return f"{n:02d}"
    return str(n)


# Fuentes con sustituto métricamente compatible y libre.
FONT_MAP = {
    "Calibri": "Carlito, Calibri, sans-serif",
    "Georgia": "Gelasio, Georgia, serif",
}

# Familia genérica de respaldo para fuentes habituales (si no están instaladas,
# al menos caen en el género correcto en vez de siempre en sans-serif).
GENERIC_FAMILY = {
    "Times New Roman": "serif",
    "Cambria": "serif",
    "Garamond": "serif",
    "Book Antiqua": "serif",
    "Palatino Linotype": "serif",
    "Courier New": "monospace",
    "Consolas": "monospace",
    "Lucida Console": "monospace",
}

# Colores con nombre del resaltado de Word (<w:highlight w:val="yellow"/>).
HIGHLIGHT_COLORS = {
    "black": "#000000", "blue": "#0000FF", "cyan": "#00FFFF",
    "darkBlue": "#000080", "darkCyan": "#008080", "darkGray": "#808080",
    "darkGreen": "#008000", "darkMagenta": "#800080", "darkRed": "#800000",
    "darkYellow": "#808000", "green": "#00FF00", "lightGray": "#C0C0C0",
    "magenta": "#FF00FF", "red": "#FF0000", "white": "#FFFFFF",
    "yellow": "#FFFF00",
}

# Glifos de viñeta más habituales por carácter o por fuente de símbolos
# (Wingdings/Symbol usan code points privados); se cae a "•" si no se reconoce.
BULLET_GLYPHS = {
    "": "•", "": "▪", "": "✓", "": "➢",
    "o": "o", "•": "•", "▪": "▪", "·": "·",
    "–": "–", "−": "–", "*": "•",
}

# Interlineado por defecto (ajustable para casar con el motor de referencia).
BODY_LINE_HEIGHT = float(os.environ.get("BODY_LH", "1.0"))
CELL_LINE_HEIGHT = float(os.environ.get("CELL_LH", "1.16"))

# Un .docx no guarda páginas fijas: Word las calcula al maquetar. Para acercar
# la paginación del PDF a la de Word respetamos sus "pistas" de salto de página
# (<w:lastRenderedPageBreak/>), que Word escribe donde partió la página la última
# vez que la renderizó. Es una aproximación (puede quedar obsoleta si el .docx se
# editó sin reabrir en Word); se puede desactivar con RESPECT_PAGE_HINTS=0.
RESPECT_PAGE_HINTS = os.environ.get("RESPECT_PAGE_HINTS", "1") not in ("0", "false", "off")


def font_stack(name: Optional[str]) -> Optional[str]:
    if not name:
        return None
    if name in FONT_MAP:
        return FONT_MAP[name]
    generic = GENERIC_FAMILY.get(name, "sans-serif")
    return f"'{name}', {generic}"


# ----------------------------------------------------------------------------
# Resolución de formato de "run" (carácter)
# ----------------------------------------------------------------------------
def rpr_dict(rpr) -> dict:
    """Extrae propiedades de carácter de un <w:rPr>."""
    d: dict[str, Any] = {}
    if rpr is None:
        return d
    fonts = first(rpr, "rFonts")
    if fonts is not None:
        if attr(fonts, "ascii"):
            d["font"] = attr(fonts, "ascii")
        elif attr(fonts, "asciiTheme"):
            # fuente de tema (p.ej. minorHAnsi -> Calibri); se resuelve luego
            # contra theme1.xml. font=None pisa cualquier fuente heredada al
            # fusionar (la rFonts del run manda sobre la del estilo por defecto).
            d["font_theme"] = attr(fonts, "asciiTheme")
            d["font"] = None
    b = on(first(rpr, "b"))
    if b is not None:
        d["bold"] = b
    i = on(first(rpr, "i"))
    if i is not None:
        d["italic"] = i
    strike = on(first(rpr, "strike"))
    if strike is not None:
        d["strike"] = strike
    u = first(rpr, "u")
    if u is not None:
        d["underline"] = attr(u, "val") not in (None, "none")
    color = first(rpr, "color")
    if color is not None:
        v = attr(color, "val")
        if v and v != "auto":
            d["color"] = "#" + v
    sz = first(rpr, "sz")
    if sz is not None:
        d["size"] = float(attr(sz, "val")) / 2.0
    va = first(rpr, "vertAlign")
    if va is not None:
        d["va"] = attr(va, "val")
    hl = first(rpr, "highlight")
    if hl is not None:
        v = attr(hl, "val")
        if v and v != "none":
            d["highlight"] = HIGHLIGHT_COLORS.get(v, v)
    caps = on(first(rpr, "caps"))
    if caps is not None:
        d["caps"] = caps
    smallcaps = on(first(rpr, "smallCaps"))
    if smallcaps is not None:
        d["smallcaps"] = smallcaps
    return d


def run_css(d: dict) -> str:
    css = []
    if d.get("font"):
        css.append(f"font-family:{font_stack(d['font'])}")
    if "bold" in d:
        css.append("font-weight:" + ("bold" if d["bold"] else "normal"))
    if "italic" in d:
        css.append("font-style:" + ("italic" if d["italic"] else "normal"))
    deco = []
    if d.get("underline"):
        deco.append("underline")
    if d.get("strike"):
        deco.append("line-through")
    if deco:
        css.append("text-decoration:" + " ".join(deco))
    if d.get("color"):
        css.append("color:" + d["color"])
    if d.get("highlight"):
        css.append("background-color:" + d["highlight"])
    # caps -> mayúsculas; smallCaps -> versalitas. Word da prioridad a caps.
    if d.get("caps"):
        css.append("text-transform:uppercase")
    elif d.get("smallcaps"):
        css.append("font-variant:small-caps")
    size = d.get("size")
    va = d.get("va")
    if va in ("superscript", "subscript"):
        css.append("vertical-align:" + ("super" if va == "superscript" else "sub"))
        if size:
            size = size * 0.7
    if size:
        css.append(f"font-size:{size:.1f}pt")
    return ";".join(css)


def border_css(b) -> Optional[str]:
    """CSS de un borde OOXML (<w:top>/<w:bottom>/...)."""
    if b is None:
        return None
    val = attr(b, "val")
    if val in (None, "nil", "none"):
        return "none"
    sz = attr(b, "sz")
    width = max(float(sz) / 8.0, 0.5) if sz else 0.5
    color = attr(b, "color") or "000000"
    if color == "auto":
        color = "000000"
    return f"{width:.2f}pt solid #{color}"


class Converter:
    def __init__(self, path: str):
        self.z: Optional[zipfile.ZipFile] = zipfile.ZipFile(path)
        self._read_bytes = 0
        self.doc = self._require_xml_part("word/document.xml")
        self.styles = self._require_xml_part("word/styles.xml")
        self.rels = self._index_rels()
        self.theme_fonts = self._index_theme()
        self.def_rpr, self.def_ppr = self._doc_defaults()
        self.style_ppr: dict[str, Any] = {}  # styleId -> resolved paragraph props
        self.style_rpr = self._index_styles()
        self.num_levels = self._index_numbering()
        # Default font/size; overridden by <w:docDefaults> when present.
        self.default: dict[str, Any] = {"font": "Calibri", "color": "#000000", "size": 10.0}
        for k in ("font", "size", "color", "bold", "italic"):
            if k in self.def_rpr and self.def_rpr[k] is not None:
                self.default[k] = self.def_rpr[k]
        self._img_cache: dict[str, str] = {}
        self._pending_floats: list[str] = []
        self._list_counters: dict[str, Any] = {}  # numId -> {ilvl: current count}
        self._content_started = False

        # cabecera/pie por tipo (default / first / even) según el sectPr
        sect = self.doc.find(w("body")).find(w("sectPr"))
        self.headers = {t: self._ref_part(sect, "headerReference", t)
                        for t in ("default", "first", "even")}
        self.footers = {t: self._ref_part(sect, "footerReference", t)
                        for t in ("default", "first", "even")}
        if self.headers["default"] is None:
            self.headers["default"] = self._opt("word/header1.xml")
        if self.footers["default"] is None:
            self.footers["default"] = self._opt("word/footer1.xml")

        # primera página distinta (<w:titlePg/>) y pares/impares distintos
        self.title_pg = bool(on(first(sect, "titlePg"))) if sect is not None else False
        try:
            settings = self._xml_part("word/settings.xml")
        except KeyError:
            settings = None
        self.even_odd = (settings is not None
                         and settings.find(w("evenAndOddHeaders")) is not None)

    # -- context manager / recursos ---------------------------------------
    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()
        return False

    def close(self) -> None:
        """Close the .docx (releases the ZIP file descriptor)."""
        if self.z is not None:
            self.z.close()
            self.z = None

    # -- lectura segura del ZIP -------------------------------------------
    def _read(self, name: str) -> bytes:
        if self.z is None:
            raise RuntimeError("Converter is already closed")
        info = self.z.getinfo(name)
        if info.file_size > MAX_MEMBER_BYTES:
            raise ValueError(f"oversized member in .docx: {name}")
        self._read_bytes += info.file_size
        if self._read_bytes > MAX_TOTAL_BYTES:
            raise ValueError("uncompressed .docx exceeds the maximum allowed size")
        return self.z.read(name)

    def _xml_part(self, name: str):
        return _xml(self._read(name))

    def _require_xml_part(self, name: str):
        """Like _xml_part but raises ValueError (not KeyError) when absent."""
        try:
            return self._xml_part(name)
        except KeyError as exc:
            raise ValueError(f"required OOXML part missing from .docx: {name}") from exc

    def _opt(self, name: str):
        try:
            return self._xml_part(name)
        except KeyError:
            return None

    def _index_rels(self) -> dict:
        try:
            root = self._xml_part("word/_rels/document.xml.rels")
        except KeyError:
            return {}
        return {r.get("Id"): r.get("Target") for r in root}

    def _ref_part(self, sect, tag: str, type_: str = "default"):
        """Carga la parte (header/footer) referenciada con el type dado."""
        if sect is None:
            return None
        for ref in sect.findall(w(tag)):
            if ref.get(w("type")) == type_:
                rid = ref.get(f"{{{R}}}id")
                target = self.rels.get(rid)
                if target:
                    return self._opt("word/" + target)
        return None

    # -- tema / valores por defecto / herencia de estilos -----------------
    def _index_theme(self) -> dict:
        """{'major': 'Calibri Light', 'minor': 'Calibri'} desde theme1.xml.

        Word suele referirse a las fuentes por tema (asciiTheme="minorHAnsi")
        en vez de por nombre; el nombre real vive en el esquema de fuentes del
        tema, así que lo indexamos para resolverlas después.
        """
        theme = self._opt("word/theme/theme1.xml")
        out: dict[str, Any] = {}
        if theme is None:
            return out
        for key, tag in (("major", "majorFont"), ("minor", "minorFont")):
            fs = theme.find(".//" + f"{{{A}}}{tag}")
            if fs is not None:
                latin = fs.find(f"{{{A}}}latin")
                if latin is not None and latin.get("typeface"):
                    out[key] = latin.get("typeface")
        return out

    def _doc_metadata(self) -> dict:
        """Título/autor/asunto/palabras clave desde docProps/core.xml.

        WeasyPrint los traslada a los metadatos del PDF vía <title>/<meta>.
        """
        core = self._opt("docProps/core.xml")
        if core is None:
            return {}
        out: dict[str, Any] = {}
        for key, ns, tag in (
            ("title", DC, "title"),
            ("author", DC, "creator"),
            ("subject", DC, "subject"),
            ("description", DC, "description"),
            ("keywords", CP, "keywords"),
        ):
            el = core.find(f"{{{ns}}}{tag}")
            if el is not None and el.text and el.text.strip():
                out[key] = el.text.strip()
        return out

    def _resolve_theme_font(self, d: dict) -> None:
        """Si el dict de formato apunta a una fuente de tema, fija su nombre."""
        if not d.get("font") and d.get("font_theme"):
            key = "major" if d["font_theme"].startswith("major") else "minor"
            name = self.theme_fonts.get(key)
            if name:
                d["font"] = name

    def _doc_defaults(self):
        """Formato por defecto del documento (<w:docDefaults>): (rPr, pPr)."""
        dd = first(self.styles, "docDefaults")
        rpr_def, ppr_def = {}, {}
        if dd is not None:
            rprd = first(dd, "rPrDefault")
            if rprd is not None:
                rpr_def = rpr_dict(first(rprd, "rPr"))
                self._resolve_theme_font(rpr_def)
            pprd = first(dd, "pPrDefault")
            if pprd is not None:
                ppr_def = self._ppr_layout(first(pprd, "pPr"))
        return rpr_def, ppr_def

    def _ppr_layout(self, ppr) -> dict:
        """Propiedades de párrafo heredables (alineación, espaciado, sangría).

        Se extraen como dict para poder fusionar las del estilo (vía basedOn)
        con las propias del párrafo, igual que hace Word.
        """
        p: dict[str, Any] = {}
        if ppr is None:
            return p
        jc = first(ppr, "jc")
        if jc is not None and attr(jc, "val"):
            p["align"] = attr(jc, "val")
        sp = first(ppr, "spacing")
        if sp is not None:
            if attr(sp, "before") is not None:
                p["before"] = attr(sp, "before")
            if attr(sp, "after") is not None:
                p["after"] = attr(sp, "after")
            if attr(sp, "line") is not None and attr(sp, "lineRule") in (None, "auto"):
                p["line"] = attr(sp, "line")
        ind = first(ppr, "ind")
        if ind is not None:
            for k in ("left", "right", "hanging", "firstLine"):
                if attr(ind, k) is not None:
                    p[k] = attr(ind, k)
        return p

    def _index_styles(self) -> dict:
        """styleId -> rPr resuelto; rellena self.style_ppr con el pPr resuelto.

        Resuelve la cadena ``w:basedOn`` para que un estilo herede el formato
        (carácter y párrafo) de su padre, con protección frente a ciclos.
        """
        raw = {}
        # Estilos por defecto (w:default="1"): Word los aplica a párrafos sin
        # un pStyle explícito. Guardamos el del tipo "paragraph" y "character".
        self.default_pstyle = None
        self.default_rstyle = None
        for st in self.styles.findall(w("style")):
            sid = attr(st, "styleId")
            based = first(st, "basedOn")
            raw[sid] = {
                "rpr": rpr_dict(first(st, "rPr")),
                "ppr": self._ppr_layout(first(st, "pPr")),
                "based": attr(based, "val") if based is not None else None,
            }
            if attr(st, "default") in ("1", "true", "on"):
                stype = attr(st, "type")
                if stype == "paragraph" and self.default_pstyle is None:
                    self.default_pstyle = sid
                elif stype == "character" and self.default_rstyle is None:
                    self.default_rstyle = sid

        resolved: dict[str, Any] = {}

        def resolve(sid, seen):
            if sid in resolved:
                return resolved[sid]
            node = raw.get(sid)
            if node is None or sid in seen:
                return {"rpr": {}, "ppr": {}}
            seen = seen | {sid}
            parent = (resolve(node["based"], seen) if node["based"]
                      else {"rpr": {}, "ppr": {}})
            merged_rpr = dict(parent["rpr"]); merged_rpr.update(node["rpr"])
            merged_ppr = dict(parent["ppr"]); merged_ppr.update(node["ppr"])
            resolved[sid] = {"rpr": merged_rpr, "ppr": merged_ppr}
            return resolved[sid]

        out: dict[str, Any] = {}
        for sid in raw:
            r = resolve(sid, set())
            self._resolve_theme_font(r["rpr"])
            out[sid] = r["rpr"]
            self.style_ppr[sid] = r["ppr"]
        return out

    def _index_numbering(self) -> dict:
        """numId -> {ilvl: {fmt, text, start, left, hanging}} desde numbering.xml."""
        try:
            numx = self._xml_part("word/numbering.xml")
        except KeyError:
            return {}
        abstract = {}
        for an in numx.findall(w("abstractNum")):
            aid = an.get(w("abstractNumId"))
            levels = {}
            for lvl in an.findall(w("lvl")):
                ilvl = int(lvl.get(w("ilvl")) or 0)
                fmt = first(lvl, "numFmt")
                txt = first(lvl, "lvlText")
                start = first(lvl, "start")
                ppr = first(lvl, "pPr")
                ind = first(ppr, "ind") if ppr is not None else None
                levels[ilvl] = {
                    "fmt": attr(fmt, "val") if fmt is not None else "decimal",
                    "text": attr(txt, "val") if txt is not None else "",
                    "start": (int(attr(start, "val"))
                              if start is not None and attr(start, "val") else 1),
                    "left": attr(ind, "left") if ind is not None else None,
                    "hanging": attr(ind, "hanging") if ind is not None else None,
                }
            abstract[aid] = levels
        out: dict[str, Any] = {}
        for num in numx.findall(w("num")):
            nid = num.get(w("numId"))
            a = first(num, "abstractNumId")
            aid = attr(a, "val") if a is not None else None
            if aid in abstract:
                out[nid] = abstract[aid]
        return out

    def _bullet_glyph(self, num_id, ilvl: int) -> str:
        """Glifo de viñeta del nivel (mapeado a un equivalente Unicode)."""
        level = self.num_levels.get(num_id, {}).get(ilvl)
        text = (level or {}).get("text") or ""
        return BULLET_GLYPHS.get(text.strip(), "•") if text.strip() else "•"

    def _list_marker(self, num_id, ilvl: int) -> Optional[str]:
        """Marcador de lista numerada (p.ej. '1.', 'a)', 'IV.') o None si viñeta."""
        levels = self.num_levels.get(num_id)
        if not levels:
            return None
        level = levels.get(ilvl)
        if level is None or level["fmt"] == "bullet":
            return None
        counters = self._list_counters.setdefault(num_id, {})
        counters[ilvl] = counters.get(ilvl, level["start"] - 1) + 1
        for deeper in [k for k in list(counters) if k > ilvl]:
            del counters[deeper]  # reiniciar niveles más profundos
        text = level["text"] or ("%" + str(ilvl + 1) + ".")

        def repl(m):
            idx = int(m.group(1)) - 1  # %1 -> nivel 0
            ldef = levels.get(idx, level)
            val = counters.get(idx, ldef["start"])
            return _format_num(val, ldef["fmt"])

        return re.sub(r"%(\d)", repl, text)

    # -- runs --------------------------------------------------------------
    def render_runs(self, p, base: dict) -> str:
        """HTML de los runs de un párrafo, heredando 'base' (rPr de su estilo).

        Ignora los campos (fldChar/instrText) y su valor cacheado: p.ej. el
        campo PAGE del pie guarda un número que no debe imprimirse tal cual.
        """
        parts = []
        in_field = False
        for child in p:
            tag = etree.QName(child).localname
            if tag == "hyperlink":
                inner = self.render_runs(child, base)
                rid = child.get(f"{{{R}}}id")
                href = self.rels.get(rid) if rid else None
                if href:
                    parts.append(
                        f'<a href="{esc(href)}" '
                        f'style="color:inherit;text-decoration:underline">{inner}</a>'
                    )
                else:
                    parts.append(
                        f'<a style="color:inherit;text-decoration:underline">{inner}</a>'
                    )
            elif tag == "r":
                types = [fc.get(w("fldCharType")) for fc in child.findall(w("fldChar"))]
                if "begin" in types:
                    in_field = True
                skip = in_field or child.find(w("instrText")) is not None
                if "end" in types:
                    in_field = False
                if not skip:
                    parts.append(self._render_run(child, base))
        return "".join(parts)

    def _render_run(self, r, base: dict) -> str:
        d = dict(base)
        d.update(rpr_dict(first(r, "rPr")))
        self._resolve_theme_font(d)
        chunks = []
        images = []
        for child in r:
            tag = etree.QName(child).localname
            if tag == "drawing":
                # solo las imágenes EN LÍNEA van aquí; las flotantes (wp:anchor)
                # las gestiona el párrafo (float o bloque aparte)
                if child.find(f"{{{WP}}}inline") is not None:
                    img = self._render_drawing(child)
                    if img:
                        images.append(img)
            elif tag == "t":
                chunks.append(keep_spaces(child.text or ""))
            elif tag == "tab":
                chunks.append("    ")
            elif tag == "cr":
                chunks.append("<br>")
            elif tag == "br":
                if child.get(w("type")) != "page":  # el salto de página
                    chunks.append("<br>")            # se gestiona en el párrafo
        text = "".join(chunks)
        out = ""
        if text:
            css = run_css(d)
            out = f'<span style="{css}">{text}</span>' if css else text
        return out + "".join(images)

    def _data_uri(self, target: str) -> str:
        if target not in self._img_cache:
            ext = target.rsplit(".", 1)[-1].lower()
            mime = {"jpg": "jpeg", "jpeg": "jpeg", "png": "png", "gif": "gif",
                    "bmp": "bmp", "svg": "svg+xml"}.get(ext, ext)
            data = base64.b64encode(self._read("word/" + target)).decode()
            self._img_cache[target] = f"data:image/{mime};base64,{data}"
        return self._img_cache[target]

    def _img_html(self, drawing, style: str) -> str:
        blip = drawing.find(".//" + f"{{{A}}}blip")
        if blip is None:
            return ""
        target = self.rels.get(blip.get(f"{{{R}}}embed"))
        if not target:
            return ""
        ext = drawing.find(".//" + f"{{{WP}}}extent")  # tamaño en EMU -> pt
        dims = ""
        if ext is not None and ext.get("cx") and ext.get("cy"):
            dims = (f"width:{emu_pt(ext.get('cx')):.1f}pt;"
                    f"height:{emu_pt(ext.get('cy')):.1f}pt;")
        return f'<img src="{self._data_uri(target)}" style="{style}{dims}">'

    def _render_drawing(self, drawing) -> str:
        return self._img_html(drawing, BLOCK_IMG_STYLE)

    def _render_anchor(self, drawing):
        """Imagen flotante (wp:anchor). Devuelve (html, wraps).

        Con ajuste cuadrado/estrecho/transparente la maquetamos con ``float``
        para que el texto la rodee; con "arriba y abajo"/"ninguno" cae a bloque.
        """
        anchor = drawing.find(f"{{{WP}}}anchor")
        wraps = False
        side = "left"
        if anchor is not None:
            if (anchor.find(f"{{{WP}}}wrapSquare") is not None
                    or anchor.find(f"{{{WP}}}wrapTight") is not None
                    or anchor.find(f"{{{WP}}}wrapThrough") is not None):
                wraps = True
            ph = anchor.find(f"{{{WP}}}positionH")
            if ph is not None:
                al = ph.find(f"{{{WP}}}align")
                if al is not None and (al.text or "").strip() == "right":
                    side = "right"
        if wraps:
            margin = "0 8pt 4pt 0" if side == "left" else "0 0 4pt 8pt"
            style = f"float:{side};margin:{margin};max-width:50%;"
            return self._img_html(drawing, style), True
        return self._render_drawing(drawing), False

    # -- párrafos ----------------------------------------------------------
    def render_paragraph(self, p, in_cell: bool = False) -> str:
        ppr = first(p, "pPr")
        style_id = None
        base = dict(self.default)
        layout = dict(self.def_ppr)  # propiedades de párrafo de docDefaults
        if ppr is not None:
            ps = first(ppr, "pStyle")
            style_id = attr(ps, "val") if ps is not None else None
        # Sin pStyle explícito, Word aplica el estilo de párrafo por defecto
        # (w:default="1", normalmente "Normal").
        if style_id is None:
            style_id = self.default_pstyle
        if style_id and style_id in self.style_rpr:
            base.update(self.style_rpr[style_id])
            layout.update(self.style_ppr.get(style_id, {}))
        self._resolve_theme_font(base)
        # las propiedades propias del párrafo pisan a las heredadas del estilo
        layout.update(self._ppr_layout(ppr))

        css = []
        num_id = None
        ilvl = 0
        has_ind = any(k in layout for k in ("left", "right", "hanging", "firstLine"))
        if layout.get("align"):
            m = {"both": "justify", "center": "center", "right": "right",
                 "left": "left", "distribute": "justify"}
            css.append("text-align:" + m.get(layout["align"], "left"))
        if "before" in layout:
            css.append(f"margin-top:{tw_pt(layout['before']):.1f}pt")
        if "after" in layout:
            css.append(f"margin-bottom:{tw_pt(layout['after']):.1f}pt")
        if "line" in layout:
            css.append(f"line-height:{float(layout['line'])/240.0:.2f}")
        if layout.get("left"):
            css.append(f"margin-left:{tw_pt(layout['left']):.1f}pt")
        if layout.get("right"):
            css.append(f"margin-right:{tw_pt(layout['right']):.1f}pt")
        if layout.get("hanging"):
            css.append(f"text-indent:-{tw_pt(layout['hanging']):.1f}pt")
        elif layout.get("firstLine"):
            css.append(f"text-indent:{tw_pt(layout['firstLine']):.1f}pt")
        if ppr is not None:
            pbdr = first(ppr, "pBdr")
            if pbdr is not None:
                for side in ("top", "bottom", "left", "right"):
                    bc = border_css(first(pbdr, side))
                    if bc and bc != "none":
                        css.append(f"border-{side}:{bc}")
                        sp_attr = first(pbdr, side)
                        if attr(sp_attr, "space"):
                            css.append(f"padding-{side}:{float(attr(sp_attr,'space')):.0f}pt")
            numpr = first(ppr, "numPr")
            if numpr is not None:
                nid_el = first(numpr, "numId")
                ilvl_el = first(numpr, "ilvl")
                num_id = attr(nid_el, "val") if nid_el is not None else None
                if ilvl_el is not None and attr(ilvl_el, "val"):
                    ilvl = int(attr(ilvl_el, "val"))

        is_list = num_id is not None
        # sangría propia del nivel de lista (si el párrafo no la trae)
        level = self.num_levels.get(num_id, {}).get(ilvl) if num_id else None
        if level and not has_ind:
            if level.get("left"):
                css.append(f"margin-left:{tw_pt(level['left']):.1f}pt")
            if level.get("hanging"):
                css.append(f"text-indent:-{tw_pt(level['hanging']):.1f}pt")

        # tamaño/fuente por defecto del párrafo (para que también afecte a
        # bullets y a la altura de líneas vacías)
        fam = font_stack(base.get("font"))
        if fam:
            css.append(f"font-family:{fam}")
        css.append(f"font-size:{base.get('size',10.0):.1f}pt")
        if base.get("color"):
            css.append("color:" + base["color"])
        if base.get("bold"):
            css.append("font-weight:bold")
        css = [c for c in css if c]

        # salto de página explícito (<w:br w:type="page"/>) dentro del párrafo
        if p.find(".//" + w("br") + "[@" + w("type") + "='page']") is not None:
            css.append("break-after:page")

        # salto de sección que inicia página nueva (sectPr en el pPr, salvo el
        # "continuo"): en Word marca un límite de página, lo forzamos también.
        if ppr is not None and not in_cell:
            sectpr = first(ppr, "sectPr")
            if sectpr is not None:
                st = first(sectpr, "type")
                if st is None or attr(st, "val") != "continuous":
                    css.append("break-after:page")

        # pista de paginación de Word (<w:lastRenderedPageBreak/>): partimos la
        # página donde Word la partió, para acercar la maquetación a la suya.
        if (RESPECT_PAGE_HINTS and not in_cell and self._content_started
                and p.find(".//" + w("lastRenderedPageBreak")) is not None):
            css.append("break-before:page")

        inner = self.render_runs(p, base)

        # imágenes flotantes: con ajuste -> float dentro del párrafo (el texto
        # las rodea); sin ajuste -> bloque diferido tras el bloque actual.
        float_html = ""
        for dr in p.iter(w("drawing")):
            if dr.find(f"{{{WP}}}anchor") is not None:
                img, wraps = self._render_anchor(dr)
                if not img:
                    continue
                if wraps:
                    float_html += img
                else:
                    self._pending_floats.append(img)

        if is_list:
            marker = self._list_marker(num_id, ilvl)
            if marker is None:
                marker = self._bullet_glyph(num_id, ilvl)
            inner = esc(marker) + " " + inner
        if not inner.strip():
            inner = inner or " "
        inner = float_html + inner
        return f'<p style="{";".join(css)}">{inner}</p>'

    # -- tablas ------------------------------------------------------------
    def render_table(self, tbl) -> str:
        tblpr = first(tbl, "tblPr")
        tblw = first(tblpr, "tblW") if tblpr is not None else None
        style = ["border-collapse:collapse", "table-layout:fixed"]
        if tblw is not None and attr(tblw, "type") == "dxa":
            style.append(f"width:{tw_pt(attr(tblw,'w')):.1f}pt")
        jc = first(tblpr, "jc") if tblpr is not None else None
        if jc is not None and attr(jc, "val") == "center":
            style.append("margin-left:auto")
            style.append("margin-right:auto")
        tblbdr = first(tblpr, "tblBorders") if tblpr is not None else None

        # anchos de columna (layout fijo)
        cols = ""
        grid = first(tbl, "tblGrid")
        if grid is not None:
            cols = "<colgroup>" + "".join(
                f'<col style="width:{tw_pt(attr(gc,"w")):.1f}pt">'
                for gc in grid.findall(w("gridCol"))
            ) + "</colgroup>"

        # Estructura del cuerpo: posición de columna, gridSpan (horizontal) y
        # vMerge (vertical) de cada celda, para resolver rowspan.
        grid_rows = []
        for tr in tbl.findall(w("tr")):
            cells = []
            col = 0
            for tc in tr.findall(w("tc")):
                tcpr = first(tc, "tcPr")
                gs = first(tcpr, "gridSpan") if tcpr is not None else None
                span = int(attr(gs, "val")) if gs is not None and attr(gs, "val") else 1
                vm = first(tcpr, "vMerge") if tcpr is not None else None
                vmerge = None
                if vm is not None:
                    vmerge = "restart" if attr(vm, "val") == "restart" else "continue"
                cells.append({"tc": tc, "col": col, "span": span,
                              "vmerge": vmerge, "rowspan": 1})
                col += span
            grid_rows.append(cells)

        # rowspan: una celda "restart" absorbe las "continue" de su columna
        for ri, cells in enumerate(grid_rows):
            for cell in cells:
                if cell["vmerge"] == "restart":
                    rs = 1
                    for rj in range(ri + 1, len(grid_rows)):
                        cont = next(
                            (c for c in grid_rows[rj]
                             if c["col"] == cell["col"] and c["vmerge"] == "continue"),
                            None,
                        )
                        if cont is None:
                            break
                        rs += 1
                    cell["rowspan"] = rs

        rows = []
        for cells in grid_rows:
            out_cells = []
            for cell in cells:
                if cell["vmerge"] == "continue":
                    continue  # absorbida por la celda "restart" superior
                out_cells.append(
                    self._render_cell(cell["tc"], tblbdr, rowspan=cell["rowspan"])
                )
            rows.append("<tr>" + "".join(out_cells) + "</tr>")
        return f'<table style="{";".join(style)}">{cols}{"".join(rows)}</table>'

    def _render_cell(self, tc, tblbdr, rowspan: int = 1) -> str:
        tcpr = first(tc, "tcPr")
        css = ["vertical-align:top"]
        spanattr = ""
        tcbdr = first(tcpr, "tcBorders") if tcpr is not None else None
        for side in ("top", "bottom", "left", "right"):
            b = first(tcbdr, side) if tcbdr is not None else None
            if b is None and tblbdr is not None:
                b = first(tblbdr, side)
            bc = border_css(b)
            css.append(f"border-{side}:{bc if bc else 'none'}")
        if tcpr is not None:
            shd = first(tcpr, "shd")
            if shd is not None:
                fill = attr(shd, "fill")
                if fill and fill != "auto":
                    css.append(f"background-color:#{fill}")
            mar = first(tcpr, "tcMar")
            if mar is not None:
                for side in ("top", "bottom", "left", "right"):
                    m = first(mar, side)
                    if m is not None and attr(m, "w"):
                        css.append(f"padding-{side}:{tw_pt(attr(m,'w')):.1f}pt")
            else:
                css.append("padding:4pt 6pt")
            va = first(tcpr, "vAlign")
            if va is not None:
                vm = {"center": "middle", "bottom": "bottom"}.get(attr(va, "val"))
                if vm:
                    css[0] = "vertical-align:" + vm
            gs = first(tcpr, "gridSpan")
            if gs is not None:
                spanattr += f' colspan="{attr(gs,"val")}"'
        else:
            css.append("padding:4pt 6pt")
        if rowspan > 1:
            spanattr += f' rowspan="{rowspan}"'
        # Renderiza párrafos y tablas anidadas en su orden de aparición (una
        # tabla dentro de una celda no debe perderse).
        inner_parts = []
        for child in tc:
            tag = etree.QName(child).localname
            if tag == "p":
                inner_parts.append(self.render_paragraph(child, in_cell=True))
            elif tag == "tbl":
                inner_parts.append(self.render_table(child))
        inner = "".join(inner_parts)
        return f'<td{spanattr} style="{";".join(css)}">{inner}</td>'

    # -- cabecera / pie ----------------------------------------------------
    def _hf_div(self, root, width_cm: float, is_footer: bool, name: str) -> str:
        """Renderiza una cabecera/pie como elemento ``running(name)``."""
        if root is None:
            return ""
        p = root.find(w("p"))
        if p is None:
            return ""
        ppr = first(p, "pPr")
        border = ""
        if ppr is not None:
            pbdr = first(ppr, "pBdr")
            if pbdr is not None:
                side = "top" if is_footer else "bottom"
                bc = border_css(first(pbdr, side))
                if bc and bc != "none":
                    border = f"border-{side}:{bc};padding-{side}:3pt;"
        base = dict(self.default)
        inner = self.render_runs(p, base)
        pagenum = ""
        if is_footer:
            # campo PAGE -> contador de página alineado a la derecha
            inner = re.sub(r" {2,}", " ", inner)
            pagenum = '<span class="pageno" style="float:right"></span>'
        style = (
            f"position:running({name});width:{width_cm:.2f}cm;{border}"
            f"font-family:{font_stack('Calibri')};color:#4a4a4a;"
        )
        return f'<div id="{name}" style="{style}">{pagenum}{inner}</div>'

    @staticmethod
    def _slot(elem_name: Optional[str], where: str) -> str:
        """Regla @top-center/@bottom-center que apunta a un running element."""
        if elem_name is None:
            return ""
        return f"@{where} {{ content: element({elem_name}); }}"

    # -- documento completo ------------------------------------------------
    def build_html(self) -> str:
        body = self.doc.find(w("body"))
        sect = body.find(w("sectPr"))
        pgsz = first(sect, "pgSz")
        pgmar = first(sect, "pgMar")
        pw = float(attr(pgsz, "w")) if pgsz is not None else 11906
        ph = float(attr(pgsz, "h")) if pgsz is not None else 16838
        mt = float(attr(pgmar, "top")) if pgmar is not None else 1440
        mb = float(attr(pgmar, "bottom")) if pgmar is not None else 1440
        ml = float(attr(pgmar, "left")) if pgmar is not None else 1200
        mr = float(attr(pgmar, "right")) if pgmar is not None else 1200
        content_cm = tw_cm(pw - ml - mr)
        page_size = f"{tw_cm(pw):.2f}cm {tw_cm(ph):.2f}cm"

        blocks = []
        for child in body:
            tag = etree.QName(child).localname
            if tag == "p":
                blocks.append(self.render_paragraph(child))
            elif tag == "tbl":
                blocks.append(self.render_table(child))
            else:
                continue
            self._content_started = True  # ya hay contenido: a partir de aquí
                                          # sí valen las pistas de salto de página
            if self._pending_floats:  # imágenes flotantes "bloque" tras el bloque
                blocks.extend(self._pending_floats)
                self._pending_floats = []

        # Cabeceras/pies por tipo. Cada variante presente se emite como un
        # running element con nombre propio y se asocia a su regla @page.
        divs = []

        def emit(root, is_footer, name):
            html = self._hf_div(root, content_cm, is_footer, name)
            if html:
                divs.append(html)
                return name
            return None

        hdr = emit(self.headers["default"], False, "hdr")
        ftr = emit(self.footers["default"], True, "ftr")
        hdr_first = emit(self.headers["first"], False, "hdr_first")
        ftr_first = emit(self.footers["first"], True, "ftr_first")
        hdr_even = emit(self.headers["even"], False, "hdr_even")
        ftr_even = emit(self.footers["even"], True, "ftr_even")

        base_page = (
            f"@page {{ size: {page_size};\n"
            f"  margin: {tw_cm(mt):.2f}cm {tw_cm(mr):.2f}cm "
            f"{tw_cm(mb):.2f}cm {tw_cm(ml):.2f}cm;\n"
            f"  {self._slot(hdr, 'top-center')} {self._slot(ftr, 'bottom-center')} }}"
        )

        first_page = ""
        if self.title_pg:
            th = (self._slot(hdr_first, "top-center")
                  if hdr_first else "@top-center { content: none; }")
            tf = (self._slot(ftr_first, "bottom-center")
                  if ftr_first else "@bottom-center { content: none; }")
            first_page = f"@page :first {{ {th} {tf} }}"

        even_page = ""
        if self.even_odd and (hdr_even or ftr_even):
            even_page = (
                f"@page :left {{ {self._slot(hdr_even, 'top-center')} "
                f"{self._slot(ftr_even, 'bottom-center')} }}"
            )

        root_family = font_stack(self.default.get("font")) or "Carlito, Calibri, sans-serif"
        root_size = self.default.get("size", 10.0)
        root_color = self.default.get("color", "#000000")
        page_css = f"""
        {base_page}
        {first_page}
        {even_page}
        html {{ font-family: {root_family}; font-size: {root_size:.1f}pt;
                color: {root_color}; }}
        body {{ margin: 0; }}
        p {{ margin: 0; line-height: {BODY_LINE_HEIGHT}; }}
        table {{ margin: 6pt 0; font-size: {root_size:.1f}pt; }}
        td p {{ margin: 0; line-height: {CELL_LINE_HEIGHT}; }}
        .pageno::after {{ content: counter(page); }}
        """
        # Metadatos del documento -> <title>/<meta> que WeasyPrint vuelca al PDF.
        meta = self._doc_metadata()
        head_meta = ""
        if meta.get("title"):
            head_meta += f"<title>{esc(meta['title'])}</title>"
        for key in ("author", "description", "keywords"):
            if meta.get(key):
                head_meta += f"<meta name='{key}' content='{esc(meta[key])}'>"
        return (
            "<!DOCTYPE html><html><head><meta charset='utf-8'>"
            + head_meta + "<style>"
            + page_css + "</style></head><body>"
            + "".join(divs) + "".join(blocks)
            + "</body></html>"
        )


# Alias aceptados para cada motor de conversión.
_ENGINE_ALIASES = {
    "auto": "auto",
    "word": "word", "msword": "word",
    "libreoffice": "libreoffice", "soffice": "libreoffice", "lo": "libreoffice",
    "weasyprint": "weasyprint", "flow": "weasyprint", "python": "weasyprint",
}


def _convert_weasyprint(in_path: str, out_path: str) -> str:
    """Native flow: OOXML -> HTML -> PDF via WeasyPrint (approximate pagination)."""
    # Deferred import: the package (and build_html) can be used without
    # WeasyPrint — and its native libraries — being installed.
    from weasyprint import HTML

    with Converter(in_path) as conv:
        html = conv.build_html()

    def _render() -> None:
        HTML(string=html).write_pdf(out_path)

    if _WEASYPRINT_TIMEOUT > 0:
        with concurrent.futures.ThreadPoolExecutor(max_workers=1) as pool:
            future = pool.submit(_render)
            try:
                future.result(timeout=_WEASYPRINT_TIMEOUT)
            except concurrent.futures.TimeoutError as exc:
                raise TimeoutError(
                    f"WeasyPrint timed out after {_WEASYPRINT_TIMEOUT}s"
                ) from exc
    else:
        _render()

    return out_path


def convert(in_path: str, out_path: str, engine: str = "auto") -> str:
    """Convert ``in_path`` (.docx) to ``out_path`` (.pdf). Returns out_path.

    ``engine`` selects the layout engine:

    - ``"auto"`` (default): uses Microsoft Word or LibreOffice when available
      — faithful pagination, same page breaks as the original — and falls back
      to the lxml + WeasyPrint flow otherwise.
    - ``"word"`` / ``"libreoffice"`` / ``"weasyprint"``: forces that engine
      (raises if the chosen engine is not available).
    """
    from . import engines

    key = _ENGINE_ALIASES.get((engine or "auto").lower())
    if key is None:
        raise ValueError(
            f"unknown engine: {engine!r} (use auto/word/libreoffice/weasyprint)"
        )

    if not os.path.exists(in_path):
        raise FileNotFoundError(f"input file not found: {in_path}")
    if not zipfile.is_zipfile(in_path):
        raise ValueError(
            f"file is not a valid .docx (not a ZIP/OOXML): {in_path}"
        )

    if key == "auto":
        # Try real layout engines in order; degrade to WeasyPrint on failure.
        attempts = (
            ("word", engines.word_available, engines.convert_word),
            ("libreoffice", lambda: bool(engines.find_libreoffice()),
             engines.convert_libreoffice),
        )
        for label, ready, run in attempts:
            if ready():
                try:
                    return run(in_path, out_path)
                except Exception as exc:  # noqa: BLE001 — degrade with warning
                    sys.stderr.write(
                        f"[docx2pdf-py] engine '{label}' failed ({exc}); "
                        "trying next\n"
                    )
        return _convert_weasyprint(in_path, out_path)

    if key == "word":
        if not engines.word_available():
            raise RuntimeError(
                "Word engine requested but not available "
                "(requires Windows or macOS with Word installed)"
            )
        return engines.convert_word(in_path, out_path)

    if key == "libreoffice":
        if not engines.find_libreoffice():
            raise RuntimeError(
                "LibreOffice engine requested but 'soffice' was not found on this system"
            )
        return engines.convert_libreoffice(in_path, out_path)

    return _convert_weasyprint(in_path, out_path)
