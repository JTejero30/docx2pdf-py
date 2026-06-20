# docx2pdf-py

Conversión **fiel** de `.docx` a PDF. El motor por defecto es **solo Python**
(sin dependencias externas): lee el OOXML del documento (estilos reales: fuentes,
colores, bordes, sombreados, tablas, imágenes, cabecera/pie) y lo recrea como
HTML que **WeasyPrint** pagina a PDF.

```
.docx ──► leer OOXML (lxml) ──► HTML+CSS ──► WeasyPrint ──► PDF
```

Si en el sistema hay un **motor de maquetación real** (Microsoft Word o
LibreOffice), `convert()` lo aprovecha para obtener paginación fiel —**el mismo
contenido por página** que el documento—; si no, usa el flujo propio. Ver
[Paginación y motores de maquetación](#paginación-y-motores-de-maquetación).

## Instalación

```bash
pip install -e .            # desde el repo (modo desarrollo)
# o, una vez publicado:
# pip install docx2pdf-py
```

Dependencias: `weasyprint` y `lxml` (se instalan solas).

### Fuentes (importante para la fidelidad)

Si el documento usa **Calibri**/**Georgia** (no libres), instala sus equivalentes
métricamente compatibles **Carlito** y **Gelasio** en el sistema; WeasyPrint los
descubre vía fontconfig y los **incrusta** en el PDF. Instrucciones en
`requirements.txt`. Otras fuentes se usan si están instaladas.

## Uso

Como librería:

```python
from docx2pdf_py import convert

convert("entrada.docx", "salida.pdf")
```

Como comando:

```bash
docx2pdf-py entrada.docx salida.pdf
docx2pdf-py                      # usa el primer .docx del directorio -> output.pdf
```

## Qué reproduce

Portada, cabecera/pie (incluidas las variantes **primera página** y
**pares/impares** vía `titlePg` / `evenAndOddHeaders`) con número de página,
encabezados, párrafos con fuente/color/negrita/cursiva/alineación, **resaltado**
(`w:highlight`), **mayúsculas/versalitas** (`w:caps`/`w:smallCaps`),
**hipervínculos** (con su URL real), **listas numeradas** (`1.`, `a)`, `IV.`…
leídas de `numbering.xml`) y con **viñeta** (con el glifo del nivel mapeado a su
equivalente Unicode), tablas (bordes, sombreados, celdas combinadas horizontal
**y verticalmente**, y **tablas anidadas**), saltos de página explícitos e
**imágenes** (inline y flotantes; las flotantes con ajuste cuadrado/estrecho
**rodean el texto** mediante `float`). El tamaño de página (incl. apaisado) se
toma del `sectPr`. Los campos de Word (p. ej. `PAGE`) se interpretan, no se
vuelca su valor cacheado. Los **metadatos** del documento (título, autor, asunto,
palabras clave de `docProps/core.xml`) se trasladan a los del PDF.

Resuelve además las **fuentes de tema** (`asciiTheme`, p. ej. `minorHAnsi` →
Calibri) leyéndolas de `theme1.xml`, y la **herencia de estilos**: cada estilo
hereda el formato (carácter y párrafo) de su `basedOn`, se aplica el **estilo de
párrafo por defecto** (`w:default="1"`, normalmente *Normal*) a los párrafos sin
`pStyle` explícito, y se aplican los valores por defecto del documento
(`docDefaults`). Así, el tamaño/espaciado/negrita de un `Heading 1` definidos
solo en `styles.xml` también se respetan.

### Paginación y motores de maquetación

Un `.docx` **no guarda páginas fijas**: las calcula el motor de maquetación al
renderizar. Por eso `convert()` elige motor según lo que haya en el sistema (con
el parámetro `engine`):

| `engine`        | Paginación                                   | Requisitos |
|-----------------|----------------------------------------------|------------|
| `auto` (def.)   | la mejor disponible                          | —          |
| `word`          | **idéntica a Word** (mismo contenido/página) | Word (Windows/macOS) |
| `libreoffice`   | **fiel** a como LibreOffice renderiza         | LibreOffice (`soffice`) |
| `weasyprint`    | aproximada (flujo propio lxml + WeasyPrint)  | WeasyPrint |

En modo `auto` se prueba **Word → LibreOffice → WeasyPrint**: si hay un motor
real, el PDF tiene **el mismo contenido por página** que el documento; si no, o
si el motor real falla, se recurre al flujo propio (con un aviso por *stderr*).
LibreOffice usa su propio motor (muy parecido a Word, no garantizado idéntico).
En todos los casos, las fuentes que falten se sustituyen, igual que al abrir el
documento en una máquina sin esas fuentes.

```python
from docx2pdf_py import convert, default_engine
convert("entrada.docx", "salida.pdf")                    # auto
convert("entrada.docx", "salida.pdf", engine="libreoffice")
print(default_engine())                                   # qué usaría 'auto' aquí
```

```bash
docx2pdf-py entrada.docx salida.pdf --engine libreoffice
```

Se puede forzar la ruta de LibreOffice con la variable `SOFFICE_BIN`.

**Flujo propio (WeasyPrint).** Cuando se usa el motor `weasyprint` la paginación
es aproximada, pero se acerca lo posible a la de Word:

- Los **saltos de sección** que inician página (`sectPr` con tipo ≠ `continuous`)
  fuerzan un salto de página.
- Se respetan las **pistas de paginación de Word** (`<w:lastRenderedPageBreak/>`),
  que Word escribe donde partió la página la última vez que la renderizó. Es una
  aproximación (puede quedar obsoleta si el documento se editó sin reabrirlo en
  Word); se puede desactivar con la variable de entorno `RESPECT_PAGE_HINTS=0`.

## Limitaciones (conversor ligero, no un motor Word completo)

- **Listas numeradas**: se renderiza el formato del nivel, pero no se aplican
  reinicios/overrides explícitos (`lvlOverride`, `startOverride`).
- **Fuentes**: mapea Calibri→Carlito y Georgia→Gelasio (incl. las referidas por
  tema vía `asciiTheme`); el resto usa la fuente real si está instalada y, si no,
  cae en su familia genérica (serif/sans/monospace).
- **Tamaño por defecto** 10 pt e **interlineado** ajustados a estilo "ofimático"
  común (configurables vía variables de entorno `BODY_LH` / `CELL_LH`).
- **Imágenes flotantes**: el ajuste se aproxima con `float` (posición exacta por
  desplazamiento absoluto no se reproduce); "arriba y abajo"/"ninguno" caen a bloque.
- Fidelidad **visual alta**, no *pixel-perfect* (eso exigiría la fuente real y el
  motor de maquetación de Word).

## Seguridad

Un `.docx` es entrada potencialmente no confiable. El parseo del OOXML usa un
parser de `lxml` endurecido (sin resolución de entidades — evita XXE y *billion
laughs* — ni acceso a red) y hay topes defensivos frente a *zip bombs*.

## Desarrollo

```bash
pip install -e .[dev]   # instala también pytest y ruff
pytest                  # los tests cubren build_html (OOXML -> HTML), sin WeasyPrint
ruff check .            # linter (mismo chequeo que ejecuta CI)
```

La CI ejecuta, además, un *smoke test* de extremo a extremo
(`tests/e2e_smoke.py`) que convierte un `.docx` real a PDF con LibreOffice.

## Estructura

```
docx2pdf_py/
  __init__.py     → expone convert(), Converter, default_engine()
  converter.py    → conversor OOXML -> HTML -> PDF (flujo propio) + dispatch de motor
  engines.py      → backends Word / LibreOffice y detección del motor
  cli.py          → comando docx2pdf-py (incluye --engine)
tests/            → suite de pytest (no requiere WeasyPrint) + e2e_smoke.py
.github/workflows → CI (lint con ruff, pytest en varias versiones, e2e LibreOffice)
pyproject.toml    → metadatos, dependencias y configuración de ruff
main.py           → script de ejemplo (edita la ruta y ejecuta)
```

## Licencia

MIT — ver [LICENSE](LICENSE).
