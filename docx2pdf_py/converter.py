#!/usr/bin/env python3
"""docx -> PDF con SOLO librerías de Python, fiel al original.

Lee el OOXML del .docx (estilos reales: fuentes, colores, bordes, sombreados,
tablas, cabecera/pie) y lo recrea como HTML, que WeasyPrint maqueta y pagina a
PDF. Las fuentes Calibri/Georgia se mapean a sus equivalentes métricos libres
Carlito/Gelasio.

Uso:
    from docx2pdf_py import convert
    convert("entrada.docx", "salida.pdf")
"""
import base64
import html as _html
import os
import re
import zipfile
from typing import Optional

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

# Parser endurecido: un .docx es entrada no confiable. Desactivamos la
# resolución de entidades (XXE / "billion laughs") y el acceso a red.
_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False)

# Tope defensivo frente a "zip bombs": tamaño máximo descomprimido por miembro
# y en total (un .docx normal está muy por debajo de esto).
MAX_MEMBER_BYTES = 200 * 1024 * 1024
MAX_TOTAL_BYTES = 500 * 1024 * 1024

BLOCK_IMG_STYLE = "display:block;margin:6pt auto;max-width:100%;"


def _xml(data: bytes):
    return etree.fromstring(data, _PARSER)


def w(tag: str) -> str:
    return f"{{{W}}}{tag}"


def emu_pt(emu) -> float:
    return float(emu) / 12700.0


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
    return float(twips) / 20.0


def tw_cm(twips) -> float:
    return float(twips) / 566.929


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

# Interlineado por defecto (ajustable para casar con el motor de referencia).
BODY_LINE_HEIGHT = float(os.environ.get("BODY_LH", "1.0"))
CELL_LINE_HEIGHT = float(os.environ.get("CELL_LH", "1.16"))


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
    d = {}
    if rpr is None:
        return d
    fonts = first(rpr, "rFonts")
    if fonts is not None and attr(fonts, "ascii"):
        d["font"] = attr(fonts, "ascii")
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
        self.z = zipfile.ZipFile(path)
        self._read_bytes = 0
        self.doc = self._xml_part("word/document.xml")
        self.styles = self._xml_part("word/styles.xml")
        self.rels = self._index_rels()
        self.style_rpr = self._index_styles()
        self.num_levels = self._index_numbering()
        self.default = {"font": "Calibri", "color": "#000000", "size": 10.0}
        self._img_cache = {}
        self._pending_floats = []   # imágenes flotantes "bloque" tras el bloque
        self._list_counters = {}    # numId -> {ilvl: contador actual}

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

    def close(self):
        """Cierra el .docx (libera el descriptor del ZIP)."""
        if getattr(self, "z", None) is not None:
            self.z.close()
            self.z = None

    # -- lectura segura del ZIP -------------------------------------------
    def _read(self, name: str) -> bytes:
        info = self.z.getinfo(name)
        if info.file_size > MAX_MEMBER_BYTES:
            raise ValueError(f"miembro demasiado grande en el .docx: {name}")
        self._read_bytes += info.file_size
        if self._read_bytes > MAX_TOTAL_BYTES:
            raise ValueError("el .docx descomprimido excede el tamaño máximo permitido")
        return self.z.read(name)

    def _xml_part(self, name: str):
        return _xml(self._read(name))

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

    def _index_styles(self) -> dict:
        out = {}
        for st in self.styles.findall(w("style")):
            sid = attr(st, "styleId")
            out[sid] = rpr_dict(first(st, "rPr"))
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
        out = {}
        for num in numx.findall(w("num")):
            nid = num.get(w("numId"))
            a = first(num, "abstractNumId")
            aid = attr(a, "val") if a is not None else None
            if aid in abstract:
                out[nid] = abstract[aid]
        return out

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
        if ppr is not None:
            ps = first(ppr, "pStyle")
            style_id = attr(ps, "val") if ps is not None else None
            if style_id and style_id in self.style_rpr:
                base.update(self.style_rpr[style_id])

        css = []
        num_id = None
        ilvl = 0
        has_ind = False
        if ppr is not None:
            jc = first(ppr, "jc")
            if jc is not None:
                m = {"both": "justify", "center": "center", "right": "right", "left": "left"}
                css.append("text-align:" + m.get(attr(jc, "val"), "left"))
            sp = first(ppr, "spacing")
            if sp is not None:
                if attr(sp, "before") is not None:
                    css.append(f"margin-top:{tw_pt(attr(sp,'before')):.1f}pt")
                if attr(sp, "after") is not None:
                    css.append(f"margin-bottom:{tw_pt(attr(sp,'after')):.1f}pt")
                line = attr(sp, "line")
                if line is not None and attr(sp, "lineRule") in (None, "auto"):
                    css.append(f"line-height:{float(line)/240.0:.2f}")
            ind = first(ppr, "ind")
            if ind is not None:
                has_ind = True
                if attr(ind, "left"):
                    css.append(f"margin-left:{tw_pt(attr(ind,'left')):.1f}pt")
                if attr(ind, "right"):
                    css.append(f"margin-right:{tw_pt(attr(ind,'right')):.1f}pt")
                if attr(ind, "hanging"):
                    css.append(f"text-indent:-{tw_pt(attr(ind,'hanging')):.1f}pt")
                elif attr(ind, "firstLine"):
                    css.append(f"text-indent:{tw_pt(attr(ind,'firstLine')):.1f}pt")
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
        css.append(f"font-family:{font_stack(base.get('font'))}")
        css.append(f"font-size:{base.get('size',10.0):.1f}pt")
        if base.get("color"):
            css.append("color:" + base["color"])
        if base.get("bold"):
            css.append("font-weight:bold")
        css = [c for c in css if c]

        # salto de página explícito (<w:br w:type="page"/>) dentro del párrafo
        if p.find(".//" + w("br") + "[@" + w("type") + "='page']") is not None:
            css.append("break-after:page")

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
            inner = (esc(marker) if marker else "–") + " " + inner
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
        inner = "".join(self.render_paragraph(p, in_cell=True) for p in tc.findall(w("p")))
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

        page_css = f"""
        {base_page}
        {first_page}
        {even_page}
        html {{ font-family: Carlito, Calibri, sans-serif; font-size: 10pt;
                color: #000000; }}
        body {{ margin: 0; }}
        p {{ margin: 0; line-height: {BODY_LINE_HEIGHT}; }}
        table {{ margin: 6pt 0; font-size: 10pt; }}
        td p {{ margin: 0; line-height: {CELL_LINE_HEIGHT}; }}
        .pageno::after {{ content: counter(page); }}
        """
        return (
            "<!DOCTYPE html><html><head><meta charset='utf-8'><style>"
            + page_css + "</style></head><body>"
            + "".join(divs) + "".join(blocks)
            + "</body></html>"
        )


def convert(in_path: str, out_path: str) -> str:
    """Convierte ``in_path`` (.docx) a ``out_path`` (.pdf). Devuelve out_path."""
    # Import diferido: así se puede importar el paquete (y probar build_html)
    # sin tener WeasyPrint —y sus librerías de sistema— instalado.
    from weasyprint import HTML

    with Converter(in_path) as conv:
        html = conv.build_html()
    HTML(string=html).write_pdf(out_path)
    return out_path
