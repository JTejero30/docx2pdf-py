# Changelog

Todas las novedades destacables de este proyecto se documentan aquí.
El formato sigue, a grandes rasgos, [Keep a Changelog](https://keepachangelog.com/).

## [Sin publicar]

### Añadido
- **Resaltado de texto** (`w:highlight`): se reproduce con `background-color`,
  mapeando los colores con nombre de Word (yellow, green, cyan…).
- **Mayúsculas y versalitas** (`w:caps` / `w:smallCaps`) → `text-transform` /
  `font-variant`.
- **Glifos de viñeta**: las listas con viñeta usan el carácter del nivel
  (`lvlText`) mapeado a su equivalente Unicode (Wingdings/Symbol → `•`, `▪`, `✓`…)
  en lugar de un guion fijo.
- **Estilo de párrafo por defecto**: los párrafos sin `pStyle` explícito heredan
  el estilo marcado `w:default="1"` (normalmente *Normal*), como hace Word.
- **Metadatos del PDF**: título, autor, asunto y palabras clave se leen de
  `docProps/core.xml` y se trasladan a los metadatos del PDF.
- **Validación de entrada** en `convert()`: error claro si el archivo no existe
  o no es un ZIP/OOXML válido.
- **Marcador de tipos** `py.typed` para consumidores con *type checkers*.
- **CI**: trabajo de *lint* con ruff y *smoke test* de extremo a extremo
  (`tests/e2e_smoke.py`) que convierte un `.docx` real a PDF con LibreOffice.

### Cambiado
- `pyproject.toml`: configuración de ruff y `ruff` añadido a las dependencias de
  desarrollo; metadatos de autoría corregidos.

## [0.1.0]

- Versión inicial: conversión `.docx` → PDF solo con Python (lxml + WeasyPrint),
  con dispatch opcional a Word/LibreOffice para paginación fiel.
