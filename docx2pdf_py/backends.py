"""Built-in engine adapters implementing :class:`ConversionEngine`."""

from __future__ import annotations

import logging
from dataclasses import dataclass
from importlib.metadata import entry_points
from typing import cast

from . import engines
from .engine_protocol import ConversionEngine
from .models import ConversionOptions, ResolvedEngine
from .output import Pathish

_log = logging.getLogger(__name__)


@dataclass(frozen=True)
class WordEngine:
    name: ResolvedEngine = "word"

    def available(self) -> bool:
        return engines.word_available()

    def convert(
        self, input_path: Pathish, output_path: Pathish, options: ConversionOptions
    ) -> str:
        return engines.convert_word(
            input_path, output_path, timeout=options.native_engine_timeout
        )


@dataclass(frozen=True)
class LibreOfficeEngine:
    name: ResolvedEngine = "libreoffice"

    def available(self) -> bool:
        return bool(engines.find_libreoffice())

    def convert(
        self, input_path: Pathish, output_path: Pathish, options: ConversionOptions
    ) -> str:
        return engines.convert_libreoffice(
            input_path, output_path, timeout=options.native_engine_timeout
        )


@dataclass(frozen=True)
class WeasyPrintEngine:
    name: ResolvedEngine = "weasyprint"

    def available(self) -> bool:
        return True

    def convert(
        self, input_path: Pathish, output_path: Pathish, options: ConversionOptions
    ) -> str:
        # Deferred to avoid importing the large OOXML renderer during discovery.
        from .converter import _convert_weasyprint

        return _convert_weasyprint(input_path, output_path, options=options)


BUILTIN_ENGINES: tuple[ConversionEngine, ...] = (
    cast(ConversionEngine, WordEngine()),
    cast(ConversionEngine, LibreOfficeEngine()),
    cast(ConversionEngine, WeasyPrintEngine()),
)

_BUILTIN_NAMES = frozenset(e.name for e in BUILTIN_ENGINES)


def load_engine_registry() -> tuple[ConversionEngine, ...]:
    """Return built-in engines plus any discovered via the entry-point group.

    Third-party packages register engines under ``docx2pdf_py.engines``.  Only
    engines whose ``name`` differs from the built-ins are appended, so
    re-registering a built-in has no effect.
    """
    extra: list[ConversionEngine] = []
    for ep in entry_points(group="docx2pdf_py.engines"):
        if ep.name in _BUILTIN_NAMES:
            continue
        try:
            engine_cls = ep.load()
            extra.append(cast(ConversionEngine, engine_cls()))
        except Exception:
            _log.warning("Failed to load engine from entry point %r", ep.name, exc_info=True)
    return BUILTIN_ENGINES + tuple(extra)
