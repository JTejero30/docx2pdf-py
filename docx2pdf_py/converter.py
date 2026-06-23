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
import hashlib
import re
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any, Literal, Optional

from lxml import etree

from .exceptions import (
    ConversionError,
    ConversionTimeoutError,
)
from .formatting import (
    BULLET_GLYPHS,
    _format_num,
    border_css,
    font_stack,
    rpr_dict,
    run_css,
)
from .formatting import _to_letter as _to_letter
from .formatting import _to_roman as _to_roman
from .models import ConversionOptions, ConversionResult, Engine
from .ooxml import (
    CP,
    DC,
    WP,
    A,
    C,
    OOXMLPackage,
    R,
    attr,
    emu_pt,
    esc,
    first,
    keep_spaces,
    on,
    tw_cm,
    tw_pt,
    w,
)
from .output import Pathish, publish_pdf
from .processes import run_process

# Defensive limits against zip bombs: max decompressed size per member and
# in total (a normal .docx is well below these figures).
MAX_MEMBER_BYTES = 200 * 1024 * 1024
MAX_TOTAL_BYTES = 500 * 1024 * 1024
MAX_XML_ELEMENTS = 2_000_000

# WeasyPrint rendering timeout in seconds (0 = no timeout).
BLOCK_IMG_STYLE = "display:block;margin:6pt auto;max-width:100%;"
# inline-block para que varias imágenes en línea fluyan EN PARALELO (no apiladas)
INLINE_IMG_STYLE = "display:inline-block;vertical-align:bottom;max-width:100%;"


class Converter(OOXMLPackage):
    def __init__(
        self,
        path: Pathish,
        options: Optional[ConversionOptions] = None,
        asset_directory: Optional[Pathish] = None,
    ):
        self.options = options or ConversionOptions.from_environment()
        self.asset_directory = Path(asset_directory) if asset_directory else None
        if self.asset_directory:
            self.asset_directory.mkdir(parents=True, exist_ok=True)
        super().__init__(
            path,
            max_member_bytes=MAX_MEMBER_BYTES,
            max_total_bytes=MAX_TOTAL_BYTES,
            max_xml_elements=MAX_XML_ELEMENTS,
        )
        try:
            self.doc = self._require_xml_part("word/document.xml")
            self.styles = self._require_xml_part("word/styles.xml")
            self.rels = self._index_rels()
            self.theme_fonts = self._index_theme()
            self.def_rpr, self.def_ppr = self._doc_defaults()
            self.style_ppr: dict[str, Any] = {}  # styleId -> resolved paragraph props
            self.style_rpr = self._index_styles()
            self.num_levels = self._index_numbering()
            self.footnotes = self._opt("word/footnotes.xml")
            self.endnotes = self._opt("word/endnotes.xml")
        except Exception:
            self.close()
            raise
        # Default font/size; overridden by <w:docDefaults> when present.
        self.default: dict[str, Any] = {"font": "Calibri", "color": "#000000", "size": 10.0}
        for k in ("font", "size", "color", "bold", "italic"):
            if k in self.def_rpr and self.def_rpr[k] is not None:
                self.default[k] = self.def_rpr[k]
        self._img_cache: dict[str, str] = {}
        self._pending_floats: list[str] = []
        self._list_counters: dict[str, Any] = {}  # numId -> {ilvl: current count}
        self._content_started = False
        # relaciones activas al resolver imágenes/enlaces: por defecto las del
        # documento, pero al pintar una cabecera/pie se cambian a las SUYAS
        # (cada parte tiene su .rels y sus rId son locales -> evita colisiones).
        self._cur_rels = self.rels

        # cabecera/pie por tipo (default / first / even) según el sectPr.
        # Cada entrada es (root_xml, rels_de_esa_parte).
        sect = self.doc.find(w("body")).find(w("sectPr"))
        self.headers = {t: self._ref_part(sect, "headerReference", t)
                        for t in ("default", "first", "even")}
        self.footers = {t: self._ref_part(sect, "footerReference", t)
                        for t in ("default", "first", "even")}
        if self.headers["default"][0] is None:
            self.headers["default"] = (self._opt("word/header1.xml"),
                                       self._part_rels("word/header1.xml"))
        if self.footers["default"][0] is None:
            self.footers["default"] = (self._opt("word/footer1.xml"),
                                       self._part_rels("word/footer1.xml"))

        # primera página distinta (<w:titlePg/>) y pares/impares distintos
        self.title_pg = bool(on(first(sect, "titlePg"))) if sect is not None else False
        try:
            settings = self._xml_part("word/settings.xml")
        except KeyError:
            settings = None
        self.even_odd = (settings is not None
                         and settings.find(w("evenAndOddHeaders")) is not None)

    # -- context manager / recursos ---------------------------------------
    def __enter__(self) -> "Converter":
        return self

    def __exit__(self, *exc: Any) -> "Literal[False]":
        self.close()
        return False

    def _index_rels(self) -> dict[str, str]:
        try:
            root = self._xml_part("word/_rels/document.xml.rels")
        except KeyError:
            return {}
        return {r.get("Id"): r.get("Target") for r in root}

    def _ref_part(self, sect: Any, tag: str, type_: str = "default") -> Any:
        """Carga la parte (header/footer) del type dado como (root, sus_rels)."""
        if sect is None:
            return (None, {})
        for ref in sect.findall(w(tag)):
            if ref.get(w("type")) == type_:
                rid = ref.get(f"{{{R}}}id")
                target = self.rels.get(rid)
                if target:
                    part = self._resolve_part("word", target)
                    return (self._opt(part), self._part_rels(part))
        return (None, {})

    def _part_rels(self, part_path: str) -> dict[str, str]:
        """Relaciones propias de una parte (p.ej. word/footer2.xml -> sus rId).

        Imprescindible para cabeceras/pies: sus rId (imágenes, enlaces) son
        locales a su .rels y NO coinciden con los del documento.
        """
        if "/" in part_path:
            folder, name = part_path.rsplit("/", 1)
            rels_path = f"{folder}/_rels/{name}.rels"
        else:
            rels_path = f"_rels/{part_path}.rels"
        try:
            root = self._xml_part(rels_path)
        except KeyError:
            return {}
        return {r.get("Id"): r.get("Target") for r in root}

    # -- tema / valores por defecto / herencia de estilos -----------------
    def _index_theme(self) -> dict[str, Any]:
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

    def _doc_metadata(self) -> dict[str, str]:
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

    def _resolve_theme_font(self, d: dict[str, Any]) -> None:
        """Si el dict de formato apunta a una fuente de tema, fija su nombre."""
        if not d.get("font") and d.get("font_theme"):
            key = "major" if d["font_theme"].startswith("major") else "minor"
            name = self.theme_fonts.get(key)
            if name:
                d["font"] = name

    def _doc_defaults(self) -> tuple[dict[str, Any], dict[str, Any]]:
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

    def _ppr_layout(self, ppr: Any) -> dict[str, Any]:
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

    def _index_styles(self) -> dict[str, Any]:
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

        def resolve(sid: Any, seen: set[str]) -> dict[str, Any]:
            if sid in resolved:
                return resolved[sid]
            node = raw.get(sid)
            if node is None or sid in seen:
                return {"rpr": {}, "ppr": {}}
            seen = seen | {sid}
            parent = (resolve(node["based"], seen) if node["based"]
                      else {"rpr": {}, "ppr": {}})
            merged_rpr = {**dict(parent["rpr"]), **dict(node["rpr"] or {})}
            merged_ppr = {**dict(parent["ppr"]), **dict(node["ppr"] or {})}
            resolved[sid] = {"rpr": merged_rpr, "ppr": merged_ppr}
            return resolved[sid]

        out: dict[str, Any] = {}
        for sid in raw:
            r = resolve(sid, set())
            self._resolve_theme_font(r["rpr"])
            out[sid] = r["rpr"]
            self.style_ppr[sid] = r["ppr"]
        return out

    def _index_numbering(self) -> dict[str, Any]:
        """numId -> {ilvl: {fmt, text, start, left, hanging}} desde numbering.xml."""
        try:
            numx = self._xml_part("word/numbering.xml")
        except KeyError:
            return {}
        def parse_level(lvl: Any) -> dict[str, Any]:
            fmt = first(lvl, "numFmt")
            txt = first(lvl, "lvlText")
            start = first(lvl, "start")
            ppr = first(lvl, "pPr")
            ind = first(ppr, "ind") if ppr is not None else None
            return {
                "fmt": attr(fmt, "val") if fmt is not None else "decimal",
                "text": attr(txt, "val") if txt is not None else "",
                "start": (int(attr(start, "val"))
                          if start is not None and attr(start, "val") else 1),
                "left": attr(ind, "left") if ind is not None else None,
                "hanging": attr(ind, "hanging") if ind is not None else None,
            }

        abstract = {}
        for an in numx.findall(w("abstractNum")):
            aid = an.get(w("abstractNumId"))
            levels = {}
            for lvl in an.findall(w("lvl")):
                ilvl = int(lvl.get(w("ilvl")) or 0)
                levels[ilvl] = parse_level(lvl)
            abstract[aid] = levels
        out: dict[str, Any] = {}
        for num in numx.findall(w("num")):
            nid = num.get(w("numId"))
            a = first(num, "abstractNumId")
            aid = attr(a, "val") if a is not None else None
            if aid in abstract:
                levels = {level: dict(value) for level, value in abstract[aid].items()}
                for override in num.findall(w("lvlOverride")):
                    ilvl = int(override.get(w("ilvl")) or 0)
                    replacement = first(override, "lvl")
                    if replacement is not None:
                        levels[ilvl] = parse_level(replacement)
                    start_override = first(override, "startOverride")
                    if start_override is not None and attr(start_override, "val"):
                        levels.setdefault(ilvl, {"fmt": "decimal", "text": "", "start": 1,
                                                 "left": None, "hanging": None})
                        levels[ilvl]["start"] = int(attr(start_override, "val"))
                out[nid] = levels
        return out

    def _bullet_glyph(self, num_id: Any, ilvl: int) -> str:
        """Glifo de viñeta del nivel (mapeado a un equivalente Unicode)."""
        level = self.num_levels.get(num_id, {}).get(ilvl)
        text = (level or {}).get("text") or ""
        return BULLET_GLYPHS.get(text.strip(), "•") if text.strip() else "•"

    def _list_marker(self, num_id: Any, ilvl: int) -> Optional[str]:
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

        def repl(m: re.Match[str]) -> str:
            idx = int(m.group(1)) - 1  # %1 -> nivel 0
            ldef = levels.get(idx, level)
            val = counters.get(idx, ldef["start"])
            return _format_num(val, ldef["fmt"])

        return re.sub(r"%(\d)", repl, text)

    # -- runs --------------------------------------------------------------
    def render_runs(self, p: Any, base: dict[str, Any]) -> str:
        """HTML de los runs de un párrafo, heredando 'base' (rPr de su estilo).

        Maneja campos de Word: el código del campo nunca se imprime, pero SÍ su
        resultado (entradas de índice TOC, PAGEREF…). Los campos PAGE/NUMPAGES se
        sustituyen por contadores en vivo, para que el número de página case con
        la paginación real del PDF (no el valor cacheado).
        """
        parts = []
        field_stack: list[dict[str, Any]] = []
        for child in p:
            tag = etree.QName(child).localname
            if tag in ("ins", "moveTo", "smartTag", "sdt"):
                parts.append(self.render_runs(child, base))
            elif tag in ("del", "moveFrom"):
                continue
            elif tag == "hyperlink":
                inner = self.render_runs(child, base)
                rid = child.get(f"{{{R}}}id")
                href = self._cur_rels.get(rid) if rid else None
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
                for fc in child.findall(w("fldChar")):
                    t = fc.get(w("fldCharType"))
                    if t == "begin":
                        field_stack.append({"instr": "", "phase": "code", "skip": False})
                    elif t == "separate" and field_stack:
                        f = field_stack[-1]
                        f["phase"] = "result"
                        up = f["instr"].upper()
                        if "NUMPAGES" in up:  # nº total de páginas -> contador
                            parts.append('<span class="pagecount"></span>')
                            f["skip"] = True
                        elif re.search(r"\bPAGE\b", up):  # nº de página -> contador
                            parts.append('<span class="pageno"></span>')
                            f["skip"] = True
                    elif t == "end" and field_stack:
                        field_stack.pop()
                instr = child.find(w("instrText"))
                if instr is not None:  # el código del campo no se imprime
                    if field_stack:
                        field_stack[-1]["instr"] += (instr.text or "")
                    continue
                if field_stack:
                    f = field_stack[-1]
                    if f["phase"] == "code" or f["skip"]:
                        continue  # zona de código o resultado sustituido por contador
                parts.append(self._render_run(child, base))
            elif tag in ("oMath", "oMathPara"):
                equation = "".join(child.itertext()).strip()
                if equation:
                    parts.append(f'<span class="equation">{esc(equation)}</span>')
        return "".join(parts)

    def _render_run(self, r: Any, base: dict[str, Any]) -> str:
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
                    img = self._render_drawing(child, inline=True)
                    if img:
                        images.append(img)
                textbox = child.find(".//" + w("txbxContent"))
                if textbox is not None:
                    chunks.append(
                        '<span class="textbox">'
                        + "<br>".join(
                            self.render_runs(p, base) for p in textbox.findall(w("p"))
                        )
                        + "</span>"
                    )
                if child.find(".//" + f"{{{C}}}chart") is not None and not images:
                    chunks.append('<span class="unsupported chart">[Chart]</span>')
            elif tag == "pict":
                textbox = child.find(".//" + w("txbxContent"))
                if textbox is not None:
                    chunks.append(
                        '<span class="textbox">'
                        + "<br>".join(
                            self.render_runs(p, base) for p in textbox.findall(w("p"))
                        )
                        + "</span>"
                    )
                else:  # imagen VML (formato antiguo, frecuente en cabeceras/portadas)
                    vml = self._render_vml(child)
                    if vml:
                        images.append(vml)
            elif tag == "object":
                chunks.append('<span class="unsupported object">[Embedded object]</span>')
            elif tag == "t":
                chunks.append(keep_spaces(child.text or ""))
            elif tag in ("footnoteReference", "endnoteReference"):
                note_id = child.get(w("id")) or ""
                kind = "footnote" if tag == "footnoteReference" else "endnote"
                chunks.append(
                    f'<sup><a href="#{kind}-{esc(note_id)}">{esc(note_id)}</a></sup>'
                )
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

    def _image_src(self, target: str) -> str:
        if target not in self._img_cache:
            ext = target.rsplit(".", 1)[-1].lower()
            data = self._read(self._resolve_part("word", target))
            if self.asset_directory:
                safe_ext = ext if ext.isalnum() and len(ext) <= 8 else "bin"
                name = hashlib.sha256(target.encode()).hexdigest()[:20] + "." + safe_ext
                (self.asset_directory / name).write_bytes(data)
                self._img_cache[target] = f"assets/{name}"
            else:
                mime = {"jpg": "jpeg", "jpeg": "jpeg", "png": "png", "gif": "gif",
                        "bmp": "bmp", "svg": "svg+xml"}.get(ext, ext)
                encoded = base64.b64encode(data).decode()
                self._img_cache[target] = f"data:image/{mime};base64,{encoded}"
        return self._img_cache[target]

    def _img_html(self, drawing: Any, style: str) -> str:
        blip = drawing.find(".//" + f"{{{A}}}blip")
        if blip is None:
            return ""
        target = self._cur_rels.get(blip.get(f"{{{R}}}embed"))
        if not target:
            return ""
        ext = drawing.find(".//" + f"{{{WP}}}extent")  # tamaño en EMU -> pt
        dims = ""
        if ext is not None and ext.get("cx") and ext.get("cy"):
            dims = (f"width:{emu_pt(ext.get('cx')):.1f}pt;"
                    f"height:{emu_pt(ext.get('cy')):.1f}pt;")
        return f'<img src="{self._image_src(target)}" style="{style}{dims}">'

    def _render_drawing(self, drawing: Any, inline: bool = False) -> str:
        return self._img_html(drawing, INLINE_IMG_STYLE if inline else BLOCK_IMG_STYLE)

    def _render_vml(self, pict: Any) -> str:
        """Imagen VML (<w:pict>/<v:imagedata r:id=...>), formato antiguo."""
        for el in pict.iter():
            if etree.QName(el).localname == "imagedata":
                target = self._cur_rels.get(el.get(f"{{{R}}}id"))
                if target:
                    return f'<img src="{self._image_src(target)}" style="{INLINE_IMG_STYLE}">'
        return ""

    def _anchor_box(self, drawing: Any) -> dict[str, Any]:
        """Geometría de una imagen flotante <wp:anchor>: offsets (pt), tamaño,
        z-order y si va detrás del texto."""
        an = drawing.find(f"{{{WP}}}anchor")
        if an is None:
            return {}

        def pos(tag: str) -> tuple[Any, Any]:
            pe = an.find(f"{{{WP}}}{tag}")
            if pe is None:
                return (None, None)
            off = pe.find(f"{{{WP}}}posOffset")
            al = pe.find(f"{{{WP}}}align")
            off_pt = emu_pt(off.text) if (off is not None and off.text) else None
            return (off_pt, al.text if al is not None else None)

        off_h, al_h = pos("positionH")
        off_v, _ = pos("positionV")
        ext = an.find(f"{{{WP}}}extent")
        z = an.get("relativeHeight")
        return {
            "offH": off_h, "alH": al_h, "offV": off_v,
            "w": emu_pt(ext.get("cx")) if (ext is not None and ext.get("cx")) else None,
            "h": emu_pt(ext.get("cy")) if (ext is not None and ext.get("cy")) else None,
            "z": int(z) if (z and z.isdigit()) else 0,
            "behind": an.get("behindDoc") in ("1", "true"),
        }

    def _render_anchored(self, drawings: list[Any]) -> str:
        """Imágenes flotantes SIN ajuste, colocadas con posición ABSOLUTA y
        z-order -> pueden quedar en paralelo o SUPERPUESTAS (una sobre otra),
        como en Word. Los valores reales de relativeHeight son enormes, así que
        reasignamos z-index 0,1,2… ordenando de atrás hacia delante."""
        entries = []
        for dr in drawings:
            img = self._img_html(dr, INLINE_IMG_STYLE)
            if img:
                entries.append((self._anchor_box(dr), img))
        if not entries:
            return ""
        entries.sort(key=lambda e: (0 if e[0].get("behind") else 1, e[0].get("z", 0)))
        items, max_bottom = [], 0.0
        for z, (b, img) in enumerate(entries):
            off_v = b.get("offV") or 0.0
            st = "position:absolute;"
            if b.get("alH") == "center" or (b.get("offH") is None and b.get("alH") in (None, "center")):
                st += "left:50%;transform:translateX(-50%);"
            elif b.get("alH") == "right":
                st += "right:0;"
            else:
                st += f"left:{b.get('offH') or 0.0:.1f}pt;"
            st += f"top:{off_v:.1f}pt;z-index:{z};"
            items.append(f'<div style="{st}">{img}</div>')
            max_bottom = max(max_bottom, off_v + (b.get("h") or 0.0))
        return (f'<div style="position:relative;height:{max_bottom:.1f}pt">'
                + "".join(items) + "</div>")

    def _render_anchor(self, drawing: Any) -> tuple[str, bool]:
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
    def render_paragraph(self, p: Any, in_cell: bool = False) -> str:
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
            css.append(
                f"line-height:{float(layout['line'])/240.0*self.options.line_height_factor:.3f}"
            )
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

        # salto de página explícito (<w:br w:type="page"/>) dentro del párrafo.
        # Usamos break-before para que el contenido del párrafo aparezca en la
        # página siguiente; break-after dejaría el texto en la página actual y
        # generaría una página vacía al final que WeasyPrint puede colapsar.
        if p.find(".//" + w("br") + "[@" + w("type") + "='page']") is not None:
            css.append("break-before:page")

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
        if (self.options.respect_page_hints and not in_cell and self._content_started
                and p.find(".//" + w("lastRenderedPageBreak")) is not None):
            css.append("break-before:page")

        inner = self.render_runs(p, base)

        # imágenes flotantes: con ajuste -> float dentro del párrafo (el texto
        # las rodea); sin ajuste -> contenedor diferido con posición absoluta y
        # z-order, para que queden EN PARALELO o SUPERPUESTAS como en Word.
        float_html = ""
        block_anchors = []
        for dr in p.iter(w("drawing")):
            an = dr.find(f"{{{WP}}}anchor")
            if an is None:
                continue
            img, wraps = self._render_anchor(dr)
            if not img:
                continue
            if wraps:
                float_html += img
            elif an.find(f"{{{WP}}}wrapTopAndBottom") is not None:
                self._pending_floats.append(img)  # bloque, texto arriba/abajo
            else:
                block_anchors.append(dr)  # wrapNone -> superpuestas/en paralelo
        if block_anchors:
            container = self._render_anchored(block_anchors)
            if container:
                self._pending_floats.append(container)

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
    def render_table(self, tbl: Any) -> str:
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
        source_rows = tbl.findall(w("tr"))
        header_count = 0
        for row_index, cells in enumerate(grid_rows):
            out_cells = []
            for cell in cells:
                if cell["vmerge"] == "continue":
                    continue  # absorbida por la celda "restart" superior
                out_cells.append(
                    self._render_cell(cell["tc"], tblbdr, rowspan=cell["rowspan"])
                )
            trpr = first(source_rows[row_index], "trPr")
            repeating = first(trpr, "tblHeader") is not None if trpr is not None else False
            cant_split = first(trpr, "cantSplit") is not None if trpr is not None else False
            if repeating and row_index == header_count:
                header_count += 1
            row_style = ' style="break-inside:avoid"' if cant_split else ""
            rows.append(f"<tr{row_style}>" + "".join(out_cells) + "</tr>")
        head = f'<thead>{"".join(rows[:header_count])}</thead>' if header_count else ""
        body = f'<tbody>{"".join(rows[header_count:])}</tbody>'
        return f'<table style="{";".join(style)}">{cols}{head}{body}</table>'

    def _render_cell(self, tc: Any, tblbdr: Any, rowspan: int = 1) -> str:
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
    def _hf_div(self, root: Any, width_cm: float, is_footer: bool, name: str) -> str:
        """Renderiza una cabecera/pie COMPLETA (todos sus párrafos y tablas)
        como elemento ``running(name)``. Los campos PAGE/NUMPAGES se convierten
        en contadores en vivo dentro de render_runs, en su posición real."""
        if root is None:
            return ""
        # aislar las imágenes flotantes para que no se filtren al cuerpo
        saved_floats = self._pending_floats
        self._pending_floats = []
        blocks = []
        for child in root:
            t = etree.QName(child).localname
            if t == "p":
                blocks.append(self.render_paragraph(child))
            elif t == "tbl":
                blocks.append(self.render_table(child))
            elif t == "sdt":
                content = first(child, "sdtContent")
                for c in (content if content is not None else []):
                    ct = etree.QName(c).localname
                    if ct == "p":
                        blocks.append(self.render_paragraph(c))
                    elif ct == "tbl":
                        blocks.append(self.render_table(c))
            if self._pending_floats:
                blocks.extend(self._pending_floats)
                self._pending_floats = []
        self._pending_floats = saved_floats
        if not blocks:
            return ""
        style = (
            f"position:running({name});width:{width_cm:.2f}cm;"
            f"font-family:{font_stack('Calibri')};color:#4a4a4a;"
        )
        return f'<div id="{name}" style="{style}">{"".join(blocks)}</div>'

    @staticmethod
    def _slot(elem_name: Optional[str], where: str) -> str:
        """Regla @top-center/@bottom-center que apunta a un running element."""
        if elem_name is None:
            return ""
        return f"@{where} {{ content: element({elem_name}); }}"

    # -- documento completo ------------------------------------------------
    def _render_notes(self, root: Any, kind: str) -> str:
        if root is None:
            return ""
        entries = []
        singular = "footnote" if kind == "footnotes" else "endnote"
        for note in root.findall(w(singular)):
            note_id = note.get(w("id"))
            if note_id is None or int(note_id) < 0:
                continue
            inner = "".join(
                self.render_paragraph(p, in_cell=True) for p in note.findall(w("p"))
            )
            entries.append(
                f'<li id="{singular}-{esc(note_id)}" value="{esc(note_id)}">{inner}</li>'
            )
        if not entries:
            return ""
        return f'<section class="{kind}"><ol>{"".join(entries)}</ol></section>'

    def _iter_blocks(self, parent: Any):
        """Itera bloques (p/tbl) descendiendo en los w:sdt (controles de
        contenido que envuelven portadas, índices y otros elementos de galería
        de Word). Sin esto, esos bloques se perderían."""
        for child in parent:
            tag = etree.QName(child).localname
            if tag in ("p", "tbl"):
                yield child
            elif tag == "sdt":
                content = first(child, "sdtContent")
                if content is not None:
                    yield from self._iter_blocks(content)

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
        section_blocks: list[str] = []
        section_specs: list[Any] = []

        def finish_section(spec: Any) -> None:
            if not section_blocks:
                return
            index = len(section_specs)
            spec = spec if spec is not None else sect
            section_specs.append(spec)
            cols = first(spec, "cols")
            count = int(attr(cols, "num")) if cols is not None and attr(cols, "num") else 1
            gap = tw_pt(attr(cols, "space")) if cols is not None and attr(cols, "space") else 36
            blocks.append(
                f'<section class="doc-section section-{index}" '
                f'style="column-count: {count};column-gap: {gap:.1f}pt">'
                + "".join(section_blocks)
                + "</section>"
            )
            section_blocks.clear()

        for child in self._iter_blocks(body):
            tag = etree.QName(child).localname
            if tag == "p":
                section_blocks.append(self.render_paragraph(child))
            elif tag == "tbl":
                section_blocks.append(self.render_table(child))
            else:
                continue
            self._content_started = True  # ya hay contenido: a partir de aquí
                                          # sí valen las pistas de salto de página
            if self._pending_floats:  # imágenes flotantes "bloque" tras el bloque
                section_blocks.extend(self._pending_floats)
                self._pending_floats = []
            child_sect = first(first(child, "pPr"), "sectPr") if tag == "p" else None
            if child_sect is not None:
                finish_section(child_sect)
        finish_section(sect)
        notes = self._render_notes(self.footnotes, "footnotes")
        endnotes = self._render_notes(self.endnotes, "endnotes")

        # Cabeceras/pies por tipo. Cada variante presente se emite como un
        # running element con nombre propio y se asocia a su regla @page.
        divs = []

        def emit(item: Any, is_footer: bool, name: str) -> Optional[str]:
            root, rels = item
            prev = self._cur_rels
            self._cur_rels = rels or self.rels  # rels propias de la cabecera/pie
            try:
                html = self._hf_div(root, content_cm, is_footer, name)
            finally:
                self._cur_rels = prev
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
        named_pages = []
        for index, section_spec in enumerate(section_specs):
            section_size = first(section_spec, "pgSz")
            section_margin = first(section_spec, "pgMar")
            sw = float(attr(section_size, "w")) if section_size is not None else pw
            sh = float(attr(section_size, "h")) if section_size is not None else ph
            smt = float(attr(section_margin, "top")) if section_margin is not None else mt
            smb = float(attr(section_margin, "bottom")) if section_margin is not None else mb
            sml = float(attr(section_margin, "left")) if section_margin is not None else ml
            smr = float(attr(section_margin, "right")) if section_margin is not None else mr
            named_pages.append(
                f"@page section-{index} {{ size: {tw_cm(sw):.2f}cm {tw_cm(sh):.2f}cm; "
                f"margin: {tw_cm(smt):.2f}cm {tw_cm(smr):.2f}cm "
                f"{tw_cm(smb):.2f}cm {tw_cm(sml):.2f}cm; "
                f"{self._slot(hdr, 'top-center')} {self._slot(ftr, 'bottom-center')} }} "
                f".section-{index} {{ page: section-{index}; }}"
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
        {"".join(named_pages)}
        {first_page}
        {even_page}
        html {{ font-family: {root_family}; font-size: {root_size:.1f}pt;
                color: {root_color}; }}
        body {{ margin: 0; }}
        .doc-section {{ column-fill: auto; }}
        p {{ margin: 0; line-height: {self.options.body_line_height}; }}
        table {{ margin: 6pt 0; font-size: {root_size:.1f}pt; }}
        thead {{ display: table-header-group; }}
        tr {{ break-inside: auto; }}
        td p {{ margin: 0; line-height: {self.options.cell_line_height}; }}
        .footnotes, .endnotes {{ column-span: all; border-top: 0.5pt solid #777;
                                margin-top: 8pt; font-size: 0.85em; }}
        .textbox {{ display: inline-block; border: 0.5pt solid #999; padding: 2pt; }}
        .equation {{ font-family: serif; font-style: italic; }}
        .pageno::after {{ content: counter(page); }}
        .pagecount::after {{ content: counter(pages); }}
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
            + "".join(divs) + '<main class="document">' + "".join(blocks)
            + notes + endnotes + "</main>"
            + "</body></html>"
        )


def _convert_weasyprint(
    in_path: Pathish, out_path: Pathish, options: Optional[ConversionOptions] = None
) -> str:
    """Native flow: OOXML -> HTML -> PDF via WeasyPrint (approximate pagination)."""
    # Rendering runs outside this process so a timeout can terminate native
    # WeasyPrint/Pango work rather than leaving a background thread alive.
    options = options or ConversionOptions.from_environment()
    with tempfile.TemporaryDirectory() as tmp:
        assets = Path(tmp) / "assets"
        with Converter(in_path, options=options, asset_directory=assets) as conv:
            html = conv.build_html()
        html_path = Path(tmp) / "document.html"
        rendered_path = Path(tmp) / "document.pdf"
        html_path.write_text(html, encoding="utf-8")
        command = [sys.executable, "-m", "docx2pdf_py._weasy_worker",
                   str(html_path), str(rendered_path)]
        try:
            process = run_process(command, timeout=options.weasyprint_timeout or None)
        except subprocess.TimeoutExpired as exc:
            raise ConversionTimeoutError(
                f"WeasyPrint timed out after {options.weasyprint_timeout}s"
            ) from exc
        if process.returncode != 0:
            detail = (process.stderr or process.stdout).decode(errors="replace").strip()
            raise ConversionError(
                "WeasyPrint failed to render the document" + (f": {detail}" if detail else "")
            )
        return publish_pdf(rendered_path, out_path)


def convert_detailed(
    in_path: Pathish,
    out_path: Pathish,
    engine: Engine = "auto",
    options: Optional[ConversionOptions] = None,
) -> ConversionResult:
    """Compatibility import for callers using ``docx2pdf_py.converter``."""
    from .api import convert_detailed as _convert_detailed

    return _convert_detailed(in_path, out_path, engine=engine, options=options)


def convert(
    in_path: Pathish,
    out_path: Pathish,
    engine: Engine = "auto",
    options: Optional[ConversionOptions] = None,
) -> str:
    """Compatibility import for callers using ``docx2pdf_py.converter``."""
    from .api import convert as _convert

    return _convert(in_path, out_path, engine=engine, options=options)
