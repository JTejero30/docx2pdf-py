"""Public conversion orchestration, diagnostics, and batch operations."""

from __future__ import annotations

import os
import time
import zipfile
from collections.abc import Callable, Sequence
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from threading import Event
from typing import cast

from .backends import load_engine_registry
from .engine_protocol import ConversionEngine
from .exceptions import EngineUnavailableError, InvalidDocumentError
from .models import (
    BatchItemResult,
    ConversionAttempt,
    ConversionOptions,
    ConversionResult,
)
from .output import Pathish, validate_pdf

_ENGINE_ALIASES = {
    "auto": "auto",
    "word": "word",
    "msword": "word",
    "libreoffice": "libreoffice",
    "soffice": "libreoffice",
    "lo": "libreoffice",
    "weasyprint": "weasyprint",
    "flow": "weasyprint",
    "python": "weasyprint",
}


def _result(
    input_path: Pathish,
    output_path: str,
    engine: str,
    started: float,
    warnings: list[str],
    attempts: list[ConversionAttempt],
) -> ConversionResult:
    return ConversionResult(
        path=output_path,
        engine=engine,
        warnings=tuple(warnings),
        elapsed_seconds=time.perf_counter() - started,
        attempts=tuple(attempts),
        input_bytes=os.path.getsize(input_path),
        output_bytes=os.path.getsize(output_path),
        page_count=validate_pdf(output_path),
    )


def convert_detailed(
    in_path: Pathish,
    out_path: Pathish,
    engine: str = "auto",
    options: ConversionOptions | None = None,
    *,
    engine_registry: Sequence[ConversionEngine] | None = None,
) -> ConversionResult:
    """Convert a document and return backend attempts and output diagnostics."""
    started = time.perf_counter()
    options = options or ConversionOptions.from_environment()
    registry = tuple(load_engine_registry() if engine_registry is None else engine_registry)
    by_name = {backend.name.lower(): backend for backend in registry}
    requested = (engine or "auto").lower()
    key = _ENGINE_ALIASES.get(requested, requested)
    if key != "auto" and key not in by_name:
        raise ValueError(
            f"unknown engine: {engine!r}"
        )
    if not os.path.exists(in_path):
        raise FileNotFoundError(f"input file not found: {in_path}")
    if not zipfile.is_zipfile(in_path):
        raise InvalidDocumentError(
            f"file is not a valid .docx (not a ZIP/OOXML): {in_path}"
        )

    if key == "auto":
        candidates = registry
    else:
        selected = by_name.get(key)
        if selected is None:
            raise EngineUnavailableError(f"engine '{key}' is not registered")
        candidates = (selected,)

    warnings: list[str] = []
    attempts: list[ConversionAttempt] = []
    for index, backend in enumerate(candidates):
        available = backend.available()
        if not available:
            attempts.append(ConversionAttempt(backend.name, False))
            if key != "auto" or options.fallback == "never":
                display_name = {
                    "word": "Word",
                    "libreoffice": "LibreOffice",
                    "weasyprint": "WeasyPrint",
                }.get(backend.name, backend.name)
                raise EngineUnavailableError(
                    f"{display_name} engine requested but not available"
                )
            continue

        attempt_started = time.perf_counter()
        try:
            path = backend.convert(in_path, out_path, options)
        except Exception as exc:  # noqa: BLE001 - policy decides whether to degrade
            elapsed = time.perf_counter() - attempt_started
            attempts.append(
                ConversionAttempt(backend.name, True, elapsed, str(exc))
            )
            can_fallback = (
                index < len(candidates) - 1
                and key == "auto"
                and (
                    options.fallback == "always"
                    or (
                        options.fallback == "unavailable-only"
                        and isinstance(exc, EngineUnavailableError)
                    )
                )
            )
            if not can_fallback:
                raise
            warnings.append(f"engine '{backend.name}' failed ({exc}); trying next")
            continue

        attempts.append(
            ConversionAttempt(
                backend.name, True, time.perf_counter() - attempt_started
            )
        )
        return _result(in_path, path, backend.name, started, warnings, attempts)

    raise EngineUnavailableError("no requested conversion engine is available")


def convert(
    in_path: Pathish,
    out_path: Pathish,
    engine: str = "auto",
    options: ConversionOptions | None = None,
) -> str:
    """Convert ``in_path`` to PDF and return ``out_path`` for compatibility."""
    return convert_detailed(in_path, out_path, engine=engine, options=options).path


def convert_batch(
    inputs: Sequence[Pathish],
    output_directory: Pathish,
    engine: str = "auto",
    options: ConversionOptions | None = None,
    *,
    max_workers: int = 4,
    cancel_event: Event | None = None,
    on_progress: Callable[[BatchItemResult], None] | None = None,
) -> tuple[BatchItemResult, ...]:
    """Convert multiple documents concurrently with stable, collision-safe names."""
    if max_workers < 1:
        raise ValueError("max_workers must be at least 1")
    destination = Path(output_directory)
    destination.mkdir(parents=True, exist_ok=True)
    cancel_event = cancel_event or Event()
    seen: dict[str, int] = {}
    jobs: list[tuple[Pathish, Path]] = []
    for item in inputs:
        stem = Path(item).stem
        key = stem.casefold()
        seen[key] = seen.get(key, 0) + 1
        number = seen[key]
        suffix = f"-{number}" if number > 1 else ""
        target = destination / f"{stem}{suffix}.pdf"
        while target.exists():
            number += 1
            seen[key] = number
            target = destination / f"{stem}-{number}.pdf"
        jobs.append((item, target))

    def run(job: tuple[Pathish, Path]) -> BatchItemResult:
        source, target = job
        if cancel_event.is_set():
            item = BatchItemResult(str(source), str(target), cancelled=True)
        else:
            try:
                result = convert_detailed(source, target, engine=engine, options=options)
                item = BatchItemResult(str(source), str(target), result=result)
            except Exception as exc:  # noqa: BLE001 - batch results isolate per-file failures
                item = BatchItemResult(str(source), str(target), error=str(exc))
        if on_progress is not None:
            on_progress(item)
        return item

    results: list[BatchItemResult | None] = [None] * len(jobs)
    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        futures = {pool.submit(run, job): index for index, job in enumerate(jobs)}
        for future in as_completed(futures):
            index = futures[future]
            if future.cancelled():
                source, target = jobs[index]
                results[index] = BatchItemResult(
                    str(source), str(target), cancelled=True
                )
            else:
                results[index] = future.result()
            if cancel_event.is_set():
                for pending in futures:
                    pending.cancel()
    return tuple(cast(BatchItemResult, result) for result in results)


async def convert_batch_async(
    inputs: Sequence[Pathish],
    output_directory: Pathish,
    engine: str = "auto",
    options: ConversionOptions | None = None,
    *,
    max_workers: int = 4,
    cancel_event: Event | None = None,
    on_progress: Callable[[BatchItemResult], None] | None = None,
) -> tuple[BatchItemResult, ...]:
    """Async wrapper around :func:`convert_batch` for use in async frameworks.

    Runs the thread pool in the default executor so the event loop is not
    blocked.  All parameters are forwarded unchanged to :func:`convert_batch`.
    """
    import asyncio
    import functools

    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(
        None,
        functools.partial(
            convert_batch,
            inputs,
            output_directory,
            engine,
            options,
            max_workers=max_workers,
            cancel_event=cancel_event,
            on_progress=on_progress,
        ),
    )
