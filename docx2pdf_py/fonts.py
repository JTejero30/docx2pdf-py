"""Fuentes empaquetadas con la librería.

Genera reglas @font-face que apuntan a los .ttf incluidos en el paquete, para
que WeasyPrint las use SIN depender de que estén instaladas en el sistema
(fontconfig). Funciona igual en Linux y Windows (URIs file:// correctas).
"""
from __future__ import annotations

import os
from pathlib import Path

_FONT_DIR = os.path.join(os.path.dirname(__file__), "fonts")

# Nombre de familia CSS por prefijo de fichero (camelCase -> con espacios).
_FAMILY = {
    "Carlito": "Carlito", "Caladea": "Caladea", "Gelasio": "Gelasio",
    "Lato": "Lato", "Roboto": "Roboto",
    "LiberationSans": "Liberation Sans", "LiberationSerif": "Liberation Serif",
    "LiberationMono": "Liberation Mono",
    "DejaVuSans": "DejaVu Sans", "DejaVuSansMono": "DejaVu Sans Mono",
}
# Familias que son fuentes variables (un fichero cubre todo el rango de peso).
_VARIABLE = {"Gelasio", "Roboto"}


def _parse(stem: str) -> tuple[str, str, int | str, str]:
    """('LiberationSans-Bold') -> (prefijo, familia, weight, style)."""
    prefix, _, variant = stem.partition("-")
    family = _FAMILY.get(prefix, prefix)
    v = variant.lower()
    weight: int | str = 400
    if "black" in v:
        weight = 900
    elif "bold" in v:
        weight = 700
    elif "light" in v:
        weight = 300
    if prefix in _VARIABLE:  # variable: declarar todo el rango de pesos
        weight = "100 900"
    style = "italic" if "italic" in v else "normal"
    return prefix, family, weight, style


def font_face_css() -> str:
    """Reglas @font-face para todas las fuentes empaquetadas."""
    rules = []
    try:
        files = sorted(f for f in os.listdir(_FONT_DIR) if f.lower().endswith(".ttf"))
    except OSError:
        return ""
    for fn in files:
        _, family, weight, style = _parse(fn[:-4])
        uri = Path(_FONT_DIR, fn).as_uri()  # file:// correcta en Linux y Windows
        rules.append(
            f"@font-face{{font-family:'{family}';font-weight:{weight};"
            f"font-style:{style};src:url('{uri}');}}"
        )
    return "\n".join(rules)
