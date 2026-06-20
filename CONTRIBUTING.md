# Contributing to docx2pdf-py

Thank you for your interest in contributing! This document covers the development workflow, coding conventions, and how to submit a good pull request.

## Development setup

```bash
# Clone and install in editable mode with all dev dependencies
git clone https://github.com/antoniolujanoluna/docx2pdf-py.git
cd docx2pdf-py
pip install -e ".[dev]"
```

**System dependencies** (needed only for the WeasyPrint rendering path):

| Platform | Command |
|----------|---------|
| Ubuntu/Debian | `sudo apt-get install libpango-1.0-0 libharfbuzz0b libpangoft2-1.0-0` |
| macOS | `brew install pango` |
| Windows | See the [WeasyPrint docs](https://doc.courtbouillon.org/weasyprint/stable/first_steps.html) |

LibreOffice is optional but required for the `e2e` tests:

```bash
# Ubuntu/Debian
sudo apt-get install libreoffice-writer
```

## Running the tests

```bash
# Fast unit tests (no WeasyPrint required)
python -m pytest -q

# With coverage report
python -m pytest --cov=docx2pdf_py --cov-report=term-missing

# Full end-to-end test (requires LibreOffice)
python tests/e2e_smoke.py
```

## Lint and type checks

```bash
ruff check .              # style & lint
mypy docx2pdf_py/         # type checking
bandit -r docx2pdf_py/    # security scan
```

## Code conventions

- **No magic numbers**: unit conversions use named constants (`_EMU_PER_PT`, `_TWIP_PER_PT`, `_TWIP_PER_CM`).
- **Comments**: only when the *why* is non-obvious; avoid restating what the code does.
- **Error messages**: in English. All user-facing strings must be English.
- **Type hints**: add return type annotations to new public functions. The package ships `py.typed`.
- **Tests**: every new feature or bug fix must include a test. Tests that exercise the Converter do not require WeasyPrint — build synthetic `.docx` files in memory using the `make_docx` or `make_full_docx` fixtures in `tests/conftest.py`.
- **Security**: never use `shell=True` in subprocess calls; never expand untrusted content into commands.

## Project structure

```
docx2pdf_py/
  converter.py   # OOXML -> HTML (core logic, ~1 200 LOC)
  engines.py     # Word / LibreOffice / WeasyPrint backends
  cli.py         # Command-line interface
  __init__.py    # Public API exports
tests/
  conftest.py         # Shared fixtures (make_docx, make_full_docx)
  test_converter.py   # Core OOXML parsing tests
  test_features.py    # Feature-specific tests (lists, images, headers)
  test_fidelity.py    # Fidelity tests (highlight, caps, metadata)
  test_engines.py     # Engine selection and backend tests
  test_robustness.py  # Edge cases and malformed-input tests
  test_real_fixtures.py # Tests using complete OOXML packages
  e2e_smoke.py        # End-to-end LibreOffice test
```

## Adding a new OOXML feature

1. Find the relevant OOXML spec section (ECMA-376).
2. Add the parsing logic inside `Converter` (usually in `rpr_dict`, `_ppr_layout`, `render_paragraph`, or `render_table`).
3. Emit the corresponding CSS in `run_css` or inline in the render method.
4. Add unit tests in the most relevant test file using the `make_docx` fixture.
5. Update the feature table in `README.md`.

## Submitting a pull request

1. Fork the repository and create a branch from `main`.
2. Make your changes following the conventions above.
3. Run the full test suite and ensure it passes.
4. Open a PR against `main`. The CI will run lint, type checks, security scan, unit tests (Python 3.9–3.13), macOS engine tests, and the e2e LibreOffice test.

## Reporting bugs

Please open an issue and include:
- Python version and OS
- Minimal `.docx` that reproduces the problem (or the OOXML snippet)
- Expected vs. actual output

## License

By contributing you agree that your changes will be licensed under the [MIT License](LICENSE).
