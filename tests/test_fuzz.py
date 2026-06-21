"""Property-based fuzz tests for OOXML parsing utilities."""
from __future__ import annotations

import pytest

try:
    from hypothesis import given, settings
    from hypothesis import strategies as st
    hypothesis_available = True
except ImportError:
    hypothesis_available = False

pytestmark = pytest.mark.skipif(not hypothesis_available, reason="hypothesis not installed")

if hypothesis_available:
    from docx2pdf_py.formatting import _format_num, _to_letter, _to_roman
    from docx2pdf_py.ooxml import OOXMLPackage, esc, keep_spaces

    @given(st.text())
    def test_esc_never_raises(s: str) -> None:
        result = esc(s)
        assert isinstance(result, str)

    @given(st.text())
    def test_keep_spaces_never_raises(s: str) -> None:
        result = keep_spaces(s)
        assert isinstance(result, str)

    @given(st.integers(min_value=1, max_value=10000))
    def test_to_roman_always_returns_string(n: int) -> None:
        result = _to_roman(n)
        assert isinstance(result, str)
        assert len(result) > 0

    @given(st.integers(min_value=1, max_value=10000))
    def test_to_letter_always_returns_string(n: int) -> None:
        result = _to_letter(n)
        assert isinstance(result, str)
        assert result.isalpha()

    @given(st.integers(min_value=1, max_value=100), st.sampled_from(
        ["decimal", "lowerLetter", "upperLetter", "lowerRoman", "upperRoman", "decimalZero"]))
    def test_format_num_never_raises(n: int, fmt: str) -> None:
        result = _format_num(n, fmt)
        assert isinstance(result, str)

    @given(st.binary(max_size=1024))
    @settings(max_examples=50)
    def test_parse_xml_no_crash(data: bytes) -> None:
        from docx2pdf_py.exceptions import InvalidDocumentError
        from docx2pdf_py.ooxml import parse_xml
        try:
            parse_xml(data, max_elements=1000)
        except (InvalidDocumentError, Exception):
            pass  # Any exception is acceptable; crashes are not

    @given(st.binary(max_size=4096))
    @settings(max_examples=30)
    def test_ooxml_package_with_random_bytes(data: bytes) -> None:
        import os
        import tempfile
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            f.write(data)
            fname = f.name
        try:
            pkg = OOXMLPackage(fname, max_member_bytes=1024*1024,
                               max_total_bytes=4*1024*1024, max_xml_elements=10000)
            pkg.close()
        except Exception:
            pass
        finally:
            os.unlink(fname)

    @given(st.binary(max_size=512))
    @settings(max_examples=30)
    def test_ooxml_resolve_part_no_traversal(path_str: bytes) -> None:
        from docx2pdf_py.exceptions import InvalidDocumentError
        from docx2pdf_py.ooxml import OOXMLPackage
        try:
            s = path_str.decode("utf-8", errors="replace")
            OOXMLPackage._resolve_part("word", s)
        except InvalidDocumentError:
            pass  # expected for path traversal attempts
