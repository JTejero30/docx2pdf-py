"""Built-in engine adapters implementing :class:`ConversionEngine`."""

from __future__ import annotations

import logging
import sys
from dataclasses import dataclass
from importlib.metadata import EntryPoint, entry_points
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


# Orden de preferencia de 'auto': LibreOffice -> Word -> WeasyPrint (Python puro).
# LibreOffice da paginación fiel y está disponible en servidores Linux sin Office.
BUILTIN_ENGINES: tuple[ConversionEngine, ...] = (
    cast(ConversionEngine, LibreOfficeEngine()),
    cast(ConversionEngine, WordEngine()),
    cast(ConversionEngine, WeasyPrintEngine()),
)

_BUILTIN_NAMES = frozenset(e.name for e in BUILTIN_ENGINES)
_EP_GROUP = "docx2pdf_py.engines"


def _get_group_eps() -> list[EntryPoint]:
    """Return entry points for this package's engine group, Python 3.9+ safe.

    ``entry_points(group=...)`` keyword selection was added in Python 3.10.
    On 3.9 the function returns a plain dict and must be indexed by group name.
    """
    if sys.version_info < (3, 10):
        return list(entry_points().get(_EP_GROUP, []))
    return list(entry_points(group=_EP_GROUP))


def load_engine_registry() -> tuple[ConversionEngine, ...]:
    """Return built-in engines plus any discovered via the entry-point group.

    Third-party packages register engines under ``docx2pdf_py.engines``.  Only
    engines whose ``name`` differs from the built-ins are appended, so
    re-registering a built-in has no effect.
    """
    extra: list[ConversionEngine] = []
    for ep in _get_group_eps():
        if ep.name in _BUILTIN_NAMES:
            continue
        try:
            engine_cls = ep.load()
            extra.append(cast(ConversionEngine, engine_cls()))
        except Exception:
            _log.warning("Failed to load engine from entry point %r", ep.name, exc_info=True)
    return BUILTIN_ENGINES + tuple(extra)
