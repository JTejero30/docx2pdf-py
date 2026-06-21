"""Typed configuration and result models for the conversion API."""

from __future__ import annotations

import math
import os
from dataclasses import dataclass
from typing import Literal

Engine = Literal["auto", "word", "libreoffice", "weasyprint"]
ResolvedEngine = Literal["word", "libreoffice", "weasyprint"]
FallbackPolicy = Literal["always", "unavailable-only", "never"]


def _env_float(name: str, default: float) -> float:
    try:
        return float(os.environ.get(name, default))
    except ValueError as exc:
        raise ValueError(f"{name} must be a number") from exc


def _env_int(name: str, default: int) -> int:
    try:
        return int(os.environ.get(name, default))
    except ValueError as exc:
        raise ValueError(f"{name} must be an integer") from exc


def _env_bool(name: str, default: bool) -> bool:
    value = os.environ.get(name)
    if value is None:
        return default
    return value.lower() not in {"0", "false", "off", "no"}


@dataclass(frozen=True)
class ConversionOptions:
    """Per-conversion rendering and process settings."""

    weasyprint_timeout: int = 120
    native_engine_timeout: int = 120
    body_line_height: float = 1.0
    cell_line_height: float = 1.16
    respect_page_hints: bool = True
    fallback: FallbackPolicy = "always"

    def __post_init__(self) -> None:
        for name in ("weasyprint_timeout", "native_engine_timeout"):
            if getattr(self, name) < 0:
                raise ValueError(f"{name} must be zero or greater")
        for name in ("body_line_height", "cell_line_height"):
            value = getattr(self, name)
            if not math.isfinite(value) or value <= 0 or value > 10:
                raise ValueError(f"{name} must be a finite value between 0 and 10")
        if self.fallback not in {"always", "unavailable-only", "never"}:
            raise ValueError("fallback must be always, unavailable-only, or never")

    @classmethod
    def from_environment(cls) -> ConversionOptions:
        """Create options from the package's backwards-compatible environment variables."""
        return cls(
            weasyprint_timeout=_env_int("WEASYPRINT_TIMEOUT", 120),
            native_engine_timeout=_env_int("NATIVE_ENGINE_TIMEOUT", 120),
            body_line_height=_env_float("BODY_LH", 1.0),
            cell_line_height=_env_float("CELL_LH", 1.16),
            respect_page_hints=_env_bool("RESPECT_PAGE_HINTS", True),
            fallback=os.environ.get("DOCX2PDF_FALLBACK", "always"),  # type: ignore[arg-type]
        )


@dataclass(frozen=True)
class ConversionAttempt:
    """One backend considered during engine selection."""

    engine: str
    available: bool
    elapsed_seconds: float = 0.0
    error: str | None = None


@dataclass(frozen=True)
class ConversionResult:
    """Detailed outcome returned by :func:`convert_detailed`."""

    path: str
    engine: str
    warnings: tuple[str, ...] = ()
    elapsed_seconds: float = 0.0
    attempts: tuple[ConversionAttempt, ...] = ()
    input_bytes: int = 0
    output_bytes: int = 0
    page_count: int | None = None


@dataclass(frozen=True)
class BatchItemResult:
    """Outcome for one input passed to :func:`convert_batch`."""

    input_path: str
    output_path: str
    result: ConversionResult | None = None
    error: str | None = None
    cancelled: bool = False

    @property
    def succeeded(self) -> bool:
        """True when the conversion completed without error or cancellation."""
        return self.result is not None and not self.cancelled

    @property
    def failed(self) -> bool:
        """True when the conversion produced an error."""
        return self.error is not None
