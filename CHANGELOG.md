# Changelog

All notable changes to this project are documented here.
The format follows, loosely, [Keep a Changelog](https://keepachangelog.com/).

## [Unreleased]

### Added
- Structured logging via Python's built-in `logging` module (`logging.getLogger("docx2pdf_py")`)
- Full type annotations for `docx2pdf_py/converter.py`; mypy now checks the module at full strictness
- `--progress` flag for CLI batch mode: displays a live progress bar on stderr
- `--watch` flag for CLI single-file mode: re-converts automatically when the source file changes
- Property-based fuzz tests (`tests/test_fuzz.py`) using Hypothesis for OOXML parsing utilities
- Performance benchmark suite (`tests/benchmarks/test_perf.py`) using pytest-benchmark
- Golden-file snapshot tests (`tests/test_snapshots.py`) for the OOXML→HTML pipeline

## [0.2.0] — Unreleased

### Added
- **Batch CLI**: `docx2pdf-py file1.docx file2.docx --output-dir pdfs/` converts
  multiple files in one command, exposing the existing `convert_batch` API from
  the CLI.
- **`convert_batch` progress callback**: optional `on_progress` parameter calls a
  `Callable[[BatchItemResult], None]` after each item completes, enabling
  real-time logging and progress bars.
- **`convert_batch_async`**: async wrapper around `convert_batch` that runs the
  thread pool via `asyncio.get_event_loop().run_in_executor`, suitable for use
  in FastAPI and other async frameworks.
- **`BatchItemResult.succeeded` / `.failed` properties**: convenience shorthands
  replacing the verbose `result is not None and not cancelled` pattern.
- **Entry-point engine discovery**: third-party packages can now register engines
  via the `docx2pdf_py.engines` entry-point group; they are appended after the
  built-ins during `auto` discovery.
- **Plugin documentation**: README explains the entry-point protocol and the
  `ConversionEngine` interface required.
- **English documentation**: README, CHANGELOG, and module docstrings translated
  to English for broader PyPI audience.
- **Python 3.10 in CI matrix**: closes the gap between 3.9 and 3.11.
- **Release workflow depends on CI**: `release.yml` now requires `lint` and
  `test` to pass before publishing, preventing broken tags from reaching PyPI.

### Changed
- `weasyprint` dependency upper bound relaxed from `<72` to `<80` to avoid
  unnecessary breakage when new WeasyPrint releases ship.
- `pyproject.toml` version bumped to `0.2.0`.
- Coverage minimum target raised from 80 % to 85 %.

## [0.1.0]

### Added
- Extensible engine protocol, explicit fallback policies, and a batch
  conversion API with bounded concurrency, cancellation, and collision-safe
  output names.
- Per-conversion diagnostics: attempts, errors, timings, sizes, and page count.
- Built-in flow support for numbering overrides, footnotes/endnotes, tracked
  changes, text boxes, equations, multi-column layouts, per-section geometry,
  repeatable table headers, and indivisible rows.
- End-to-end corpus with expected per-page text and raster check.
- XML element limits, PDF structural validation, and process-tree termination.
- `convert_detailed()` API with the engine actually used, fallback warnings,
  and typed per-conversion options via `ConversionOptions`.
- Public exception hierarchy for invalid documents, unavailable engines,
  conversion errors, and timeouts.
- Atomic output publishing and PDF validation for all engines.
- Terminable processes for WeasyPrint and Word automation on Windows.
- CLI, packaging, timeout, fallback, and output-preservation tests.
- Dependency automation, wheel validation, and PyPI publishing with federated
  identity in CI.
- **Text highlighting** (`w:highlight`): reproduced with `background-color`,
  mapping Word's named colours (yellow, green, cyan…).
- **Caps / small-caps** (`w:caps` / `w:smallCaps`) → `text-transform` /
  `font-variant`.
- **Bullet glyphs**: bulleted lists use the level character from `lvlText`
  mapped to its Unicode equivalent (Wingdings/Symbol → `•`, `▪`, `✓`…)
  instead of a hard-coded dash.
- **Default paragraph style**: paragraphs without an explicit `pStyle` inherit
  the style marked `w:default="1"` (normally *Normal*), as Word does.
- **PDF metadata**: title, author, subject, and keywords are read from
  `docProps/core.xml` and transferred to the PDF.
- **Input validation** in `convert()`: clear error if the file does not exist
  or is not a valid ZIP/OOXML document.
- **`py.typed` marker** for type-checker consumers.
- **CI**: `lint` job with ruff and an end-to-end smoke test (`tests/e2e_smoke.py`)
  that converts a real `.docx` to PDF with LibreOffice.

### Changed
- WeasyPrint rendering extracts images to local temporary resources to avoid
  inflating HTML and duplicating memory via base64.
- Monolithic converter split into api, backends, ooxml, formatting, processes,
  and output modules.
- `pyproject.toml`: ruff configuration added; authorship metadata corrected.
