# docx2pdf-py

**Faithful** `.docx` to PDF conversion. The default engine is **pure Python**
(no external dependencies): reads the OOXML from the document (real styles:
fonts, colours, borders, shading, tables, images, headers/footers) and
recreates it as HTML that **WeasyPrint** lays out and paginates into a PDF.

```
.docx ──► read OOXML (lxml) ──► HTML+CSS ──► WeasyPrint ──► PDF
```

If a **real layout engine** (Microsoft Word or LibreOffice) is available on the
system, `convert()` uses it to obtain faithful pagination — **the same content
per page** as the document; otherwise it falls back to the built-in flow.
See [Pagination and layout engines](#pagination-and-layout-engines).

## Installation

```bash
pip install -e .            # from the repo (development mode)
# or, once published:
# pip install docx2pdf-py
```

Dependencies: `weasyprint` and `lxml` (installed automatically).

### Fonts (important for fidelity)

If the document uses **Calibri** / **Georgia** (non-free), install their
metrically compatible free equivalents **Carlito** and **Gelasio** on the
system; WeasyPrint discovers them via fontconfig and **embeds** them in the
PDF. See `requirements.txt` for instructions. Other fonts are used if
installed.

## Usage

As a library:

```python
from docx2pdf_py import convert

convert("input.docx", "output.pdf")
```

To know which engine was actually used and surface recoverable failures from
`auto`, use the detailed API:

```python
from docx2pdf_py import ConversionOptions, convert_detailed

result = convert_detailed(
    "input.docx",
    "output.pdf",
    options=ConversionOptions(weasyprint_timeout=60, respect_page_hints=False),
)
print(result.engine, result.warnings)
print(result.page_count, result.elapsed_seconds, result.attempts)
```

The fallback policy can be `always` (default), `unavailable-only`, or `never`.
For batch jobs there is a concurrent, bounded, and cancellable API:

```python
from docx2pdf_py import convert_batch

items = convert_batch(
    ["a.docx", "b.docx"],
    "pdfs",
    max_workers=2,
    on_progress=lambda r: print(r.input_path, "done" if r.succeeded else r.error),
)
for item in items:
    print(item.input_path, item.result or item.error)
```

An async wrapper is also available for use in async codebases (FastAPI, etc.):

```python
from docx2pdf_py import convert_batch_async

items = await convert_batch_async(["a.docx", "b.docx"], "pdfs")
```

As a command:

```bash
# Single file
docx2pdf-py input.docx output.pdf

# Auto-discover the first .docx in the current directory
docx2pdf-py

# Batch — convert several files into a directory
docx2pdf-py file1.docx file2.docx --output-dir pdfs/
docx2pdf-py *.docx --output-dir pdfs/

# Show a live progress bar while converting a batch
docx2pdf-py *.docx --output-dir pdfs/ --progress

# Watch mode — re-convert automatically whenever the source file changes
docx2pdf-py input.docx output.pdf --watch
```

## What is reproduced

Cover page, headers/footers (including the **first-page** and **odd/even**
variants via `titlePg` / `evenAndOddHeaders`) with page numbers, headings,
paragraphs with font/colour/bold/italic/alignment, **highlighting**
(`w:highlight`), **caps/small-caps** (`w:caps` / `w:smallCaps`),
**hyperlinks** (with the real URL), **numbered lists** (`1.`, `a)`, `IV.` …
read from `numbering.xml`) and **bulleted lists** (with the level glyph mapped
to its Unicode equivalent), tables (borders, shading, horizontally and
**vertically merged cells**, and **nested tables**), explicit page breaks, and
**images** (inline and floating; floating images with square/tight wrapping
**flow around text** via `float`). The page size (including landscape) is
taken from `sectPr`. Word fields (e.g. `PAGE`) are interpreted, not dumped
from cache. Document **metadata** (title, author, subject, keywords from
`docProps/core.xml`) is transferred to the PDF metadata.
Also reproduced: numbering overrides and restarts, footnotes and endnotes,
tracked-change inserted text, text boxes, text equations, multi-column
documents, per-section geometry, and repeatable table header rows.

Theme fonts (`asciiTheme`, e.g. `minorHAnsi` → Calibri) are resolved from
`theme1.xml`, and **style inheritance** is fully applied: each style inherits
formatting (character and paragraph) from its `basedOn`, the **default
paragraph style** (`w:default="1"`, normally *Normal*) is applied to
paragraphs without an explicit `pStyle`, and document default values
(`docDefaults`) are honoured. This means a `Heading 1` size/spacing/bold
defined only in `styles.xml` is also respected.

### Pagination and layout engines

A `.docx` **does not store fixed pages**: they are computed by the layout
engine when rendering. `convert()` therefore selects an engine based on what
is available (controllable via the `engine` parameter):

| `engine`        | Pagination                                      | Requirements |
|-----------------|-------------------------------------------------|--------------|
| `auto` (def.)   | best available                                  | —            |
| `word`          | **identical to Word** (same content per page)   | Word (Windows/macOS) |
| `libreoffice`   | **faithful** to LibreOffice rendering           | LibreOffice (`soffice`) |
| `weasyprint`    | approximate (built-in lxml + WeasyPrint flow)   | WeasyPrint   |

In `auto` mode **Word → LibreOffice → WeasyPrint** is tried in order: if a
real engine is found, the PDF has **the same content per page** as the
document; otherwise (or if the real engine fails) the built-in flow is used
(a warning is printed to *stderr*). LibreOffice uses its own engine (close to
Word but not guaranteed identical). In all cases missing fonts are substituted,
just as when opening the document on a machine without those fonts.

```python
from docx2pdf_py import convert, default_engine
convert("input.docx", "output.pdf")                    # auto
convert("input.docx", "output.pdf", engine="libreoffice")
print(default_engine())                                  # which engine 'auto' would pick here
```

```bash
docx2pdf-py input.docx output.pdf --engine libreoffice
docx2pdf-py input.docx output.pdf --fallback unavailable-only
```

The LibreOffice binary path can be overridden with the `SOFFICE_BIN`
environment variable.

**Built-in flow (WeasyPrint).** When the `weasyprint` engine is used,
pagination is approximate but is brought as close as possible to Word's:

- **Section breaks** that start a new page (`sectPr` with type ≠ `continuous`)
  force a page break.
- **Word page hints** (`<w:lastRenderedPageBreak/>`) are respected — Word
  writes these where it last broke the page. They can be stale if the document
  was edited without reopening in Word; disable with `RESPECT_PAGE_HINTS=0`.

## Limitations (a lightweight converter, not a full Word engine)

- **Charts and embedded objects**: when no image preview exists a placeholder
  is kept in the output, not the editable chart.
- **Tracked changes**: inserted text is shown and deleted text is hidden; review
  balloons, authors, and dates are not reproduced.
- **Fonts**: Calibri→Carlito and Georgia→Gelasio are mapped (including those
  referenced by theme via `asciiTheme`); other fonts are used if installed and
  fall back to their generic family (serif/sans/monospace) if not.
- **Default size** 10 pt and **line-height** tuned to common office style
  (configurable via `BODY_LH` / `CELL_LH` environment variables).
- **Floating images**: wrapping is approximated with `float` (exact absolute
  offset positioning is not reproduced); "top and bottom" / "none" fall back to
  block layout.
- Visual fidelity is **high**, not *pixel-perfect* (that would require the real
  font and Word's layout engine).

## Security

A `.docx` is potentially untrusted input. OOXML parsing uses a hardened `lxml`
parser (no entity resolution — prevents XXE and *billion-laughs* attacks — and
no network access) with defensive caps against *zip bombs*. XML element count
is also limited and OOXML relationship targets are normalised. Engines first
write a temporary PDF, validate its header, then atomically replace the output.
Word, LibreOffice, and WeasyPrint are all run with configurable timeouts.

## Plugin engines

Third-party packages can register custom engines via the
`docx2pdf_py.engines` entry-point group:

```toml
# In your package's pyproject.toml:
[project.entry-points."docx2pdf_py.engines"]
my-engine = "my_package.engine:MyEngine"
```

`MyEngine` must satisfy the `ConversionEngine` protocol (a `name` attribute,
`available() -> bool`, and `convert(in, out, options) -> str`). Registered
engines are appended after the built-in ones during `auto` discovery.

## Development

```bash
pip install -e .[dev]   # also installs pytest, ruff, hypothesis, pytest-benchmark
pytest                  # unit tests — cover OOXML→HTML; no WeasyPrint required
pytest tests/test_fuzz.py           # property-based fuzz tests (requires hypothesis)
pytest tests/benchmarks/ --benchmark-only   # performance benchmarks
pytest tests/test_snapshots.py --update-snapshots   # regenerate HTML golden files
ruff check .            # linter (same check that CI runs)
mypy docx2pdf_py        # type checking (all 15 modules, zero errors)
python -m build         # sdist + wheel
```

CI also runs an end-to-end smoke test (`tests/e2e_smoke.py`) that converts a
real `.docx` to PDF with LibreOffice.

Structured logging is emitted on the `docx2pdf_py` logger at `DEBUG`/`INFO`/`WARNING`
levels, so you can see engine selection and timing by enabling it in your app:

```python
import logging
logging.basicConfig(level=logging.DEBUG)
```

## Structure

```
docx2pdf_py/
  __init__.py        → exposes convert(), Converter, default_engine()
  converter.py       → OOXML document traversal and HTML/CSS assembly
  api.py             → engine selection, fallback, diagnostics, and batch
  backends.py        → engine-protocol adapters + entry-point discovery
  engine_protocol.py → extensible interface for external engines
  ooxml.py           → secure package reading and OOXML utilities
  formatting.py      → numbering, fonts, text properties, and CSS
  engines.py         → Word / LibreOffice backends and engine detection
  models.py          → typed options and detailed result objects
  output.py          → PDF validation and atomic publishing
  exceptions.py      → public error hierarchy
  _*_worker.py       → isolated, terminable processes for Word/WeasyPrint
  processes.py       → process-tree execution and termination
  cli.py             → docx2pdf-py command (single file and batch)
tests/               → pytest suite (no WeasyPrint required) + e2e scripts
.github/workflows    → CI (ruff lint, pytest on multiple versions, e2e LibreOffice)
pyproject.toml       → metadata, dependencies, and ruff/mypy configuration
main.py              → example script (edit the path and run)
```

## License

MIT — see [LICENSE](LICENSE).
