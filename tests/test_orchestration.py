"""Engine protocol, fallback policy, diagnostics, and batch conversion tests."""

from __future__ import annotations

import asyncio
import shutil
from dataclasses import dataclass
from threading import Event

import pytest

from docx2pdf_py import BatchItemResult, ConversionOptions, convert_batch
from docx2pdf_py import converter as C
from docx2pdf_py.api import convert_batch_async, convert_detailed
from tests.conftest import FAKE_PDF, document


@dataclass
class FakeEngine:
    name: str
    is_available: bool = True
    failure: Exception | None = None

    def available(self) -> bool:
        return self.is_available

    def convert(self, input_path, output_path, options):
        if self.failure:
            raise self.failure
        with open(output_path, "wb") as stream:
            stream.write(FAKE_PDF)
        return str(output_path)


def test_engine_protocol_diagnostics_and_fallback(make_docx, tmp_path):
    source = make_docx(document("<w:p/>"))
    output = tmp_path / "output.pdf"
    registry = (
        FakeEngine("word", failure=RuntimeError("word failed")),
        FakeEngine("libreoffice"),
    )

    result = convert_detailed(source, output, engine_registry=registry)

    assert result.engine == "libreoffice"
    assert result.page_count == 1
    assert result.input_bytes > 0 and result.output_bytes == len(FAKE_PDF)
    assert result.elapsed_seconds >= 0
    assert [attempt.engine for attempt in result.attempts] == ["word", "libreoffice"]
    assert result.attempts[0].error == "word failed"


@pytest.mark.parametrize("policy", ["never", "unavailable-only"])
def test_strict_fallback_policies_propagate_conversion_failures(
    make_docx, tmp_path, policy
):
    source = make_docx(document("<w:p/>"))
    registry = (
        FakeEngine("word", failure=RuntimeError("broken")),
        FakeEngine("weasyprint"),
    )
    with pytest.raises(RuntimeError, match="broken"):
        convert_detailed(
            source,
            tmp_path / "output.pdf",
            options=ConversionOptions(fallback=policy),
            engine_registry=registry,
        )


def test_unavailable_only_skips_unavailable_engine(make_docx, tmp_path):
    source = make_docx(document("<w:p/>"))
    registry = (FakeEngine("word", is_available=False), FakeEngine("weasyprint"))
    result = convert_detailed(
        source,
        tmp_path / "output.pdf",
        options=ConversionOptions(fallback="unavailable-only"),
        engine_registry=registry,
    )
    assert result.engine == "weasyprint"
    assert not result.attempts[0].available


def test_never_policy_does_not_skip_unavailable_preferred_engine(make_docx, tmp_path):
    source = make_docx(document("<w:p/>"))
    registry = (FakeEngine("word", is_available=False), FakeEngine("weasyprint"))
    with pytest.raises(RuntimeError, match="not available"):
        convert_detailed(
            source,
            tmp_path / "output.pdf",
            options=ConversionOptions(fallback="never"),
            engine_registry=registry,
        )


def test_explicit_third_party_engine_name(make_docx, tmp_path):
    source = make_docx(document("<w:p/>"))
    plugin = FakeEngine("custom")
    result = convert_detailed(
        source,
        tmp_path / "output.pdf",
        engine="custom",
        engine_registry=(plugin,),
    )
    assert result.engine == "custom"


def test_last_backend_failure_is_not_masked(make_docx, tmp_path):
    source = make_docx(document("<w:p/>"))
    registry = (FakeEngine("weasyprint", failure=RuntimeError("render failed")),)
    with pytest.raises(RuntimeError, match="render failed"):
        convert_detailed(source, tmp_path / "output.pdf", engine_registry=registry)


def test_batch_conversion_collision_names_and_cancellation(
    make_docx, tmp_path, monkeypatch
):
    original = make_docx(document("<w:p/>"))
    first = tmp_path / "a" / "same.docx"
    second = tmp_path / "b" / "same.docx"
    first.parent.mkdir()
    second.parent.mkdir()
    shutil.copyfile(original, first)
    shutil.copyfile(original, second)

    def render(_input, output, options=None):
        with open(output, "wb") as stream:
            stream.write(FAKE_PDF)
        return str(output)

    monkeypatch.setattr(C, "_convert_weasyprint", render)
    results = convert_batch(
        [first, second], tmp_path / "pdfs", engine="weasyprint", max_workers=2
    )
    assert [item.output_path for item in results] == [
        str(tmp_path / "pdfs" / "same.pdf"),
        str(tmp_path / "pdfs" / "same-2.pdf"),
    ]
    assert all(item.result and not item.error for item in results)

    cancelled = Event()
    cancelled.set()
    stopped = convert_batch([first], tmp_path / "cancelled", cancel_event=cancelled)
    assert stopped[0].cancelled


def test_batch_progress_callback_called_for_each_item(make_docx, tmp_path, monkeypatch):
    original = make_docx(document("<w:p/>"))

    def render(_input, output, options=None):
        with open(output, "wb") as stream:
            stream.write(FAKE_PDF)
        return str(output)

    monkeypatch.setattr(C, "_convert_weasyprint", render)

    seen: list[BatchItemResult] = []
    results = convert_batch(
        [original],
        tmp_path / "pdfs",
        engine="weasyprint",
        on_progress=seen.append,
    )
    assert len(seen) == 1
    assert seen[0].succeeded
    assert seen[0] is results[0]


def test_batch_item_result_succeeded_and_failed_properties(make_docx, tmp_path, monkeypatch):
    original = make_docx(document("<w:p/>"))

    def render(_input, output, options=None):
        with open(output, "wb") as stream:
            stream.write(FAKE_PDF)
        return str(output)

    monkeypatch.setattr(C, "_convert_weasyprint", render)
    results = convert_batch([original], tmp_path / "ok", engine="weasyprint")
    assert results[0].succeeded
    assert not results[0].failed

    broken = make_docx(document("<w:p/>"))
    monkeypatch.setattr(C, "_convert_weasyprint",
                        lambda *a, **kw: (_ for _ in ()).throw(RuntimeError("boom")))
    err_results = convert_batch([broken], tmp_path / "err", engine="weasyprint")
    assert not err_results[0].succeeded
    assert err_results[0].failed

    cancelled_event = Event()
    cancelled_event.set()
    cancelled_results = convert_batch([original], tmp_path / "cancelled",
                                      cancel_event=cancelled_event)
    assert not cancelled_results[0].succeeded
    assert not cancelled_results[0].failed


def test_convert_batch_async_returns_same_results(make_docx, tmp_path, monkeypatch):
    original = make_docx(document("<w:p/>"))

    def render(_input, output, options=None):
        with open(output, "wb") as stream:
            stream.write(FAKE_PDF)
        return str(output)

    monkeypatch.setattr(C, "_convert_weasyprint", render)
    results = asyncio.run(
        convert_batch_async([original], tmp_path / "async_pdfs", engine="weasyprint")
    )
    assert len(results) == 1
    assert results[0].succeeded
