"""Command-line interface: docx2pdf-py [inputs...] [output] [--output-dir DIR]"""
from __future__ import annotations

import argparse
import glob
import os
import sys
import time
from collections.abc import Sequence
from dataclasses import replace

from . import __version__
from .api import convert_batch, convert_detailed
from .exceptions import Docx2PdfError
from .models import ConversionOptions


class _ProgressBar:
    def __init__(self, total: int) -> None:
        self._total = total
        self._done = 0

    def advance(self) -> None:
        self._done += 1
        width = 30
        filled = int(width * self._done / max(self._total, 1))
        bar = "=" * filled + "-" * (width - filled)
        print(f"\r[{bar}] {self._done}/{self._total}", end="", flush=True, file=sys.stderr)

    def finish(self) -> None:
        print(file=sys.stderr)  # newline


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="docx2pdf-py",
        description=(
            "Convert .docx files to PDF using pure Python libraries.\n\n"
            "Single-file:  docx2pdf-py input.docx [output.pdf]\n"
            "Batch:        docx2pdf-py file1.docx file2.docx --output-dir pdfs/"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "inputs",
        nargs="*",
        help=(
            "path(s) to .docx file(s). "
            "Single-file mode: first arg is input, optional second arg is output path. "
            "Batch mode (--output-dir): all args are input files."
        ),
    )
    parser.add_argument(
        "--output-dir",
        metavar="DIR",
        help="batch mode: write all converted PDFs into this directory",
    )
    parser.add_argument(
        "-f", "--force",
        action="store_true",
        help="overwrite the output file if it already exists (single-file mode only)",
    )
    parser.add_argument(
        "-e", "--engine",
        default="auto",
        choices=["auto", "word", "libreoffice", "weasyprint"],
        help=(
            "layout engine to use (default: auto). "
            "'word' and 'libreoffice' produce faithful pagination; "
            "'weasyprint' uses the built-in Python flow (approximate)."
        ),
    )
    parser.add_argument(
        "-j", "--workers",
        type=int,
        default=4,
        metavar="N",
        help="maximum parallel workers for batch mode (default: 4)",
    )
    parser.add_argument(
        "-q", "--quiet",
        action="store_true",
        help="suppress all output on success",
    )
    parser.add_argument(
        "--fallback",
        choices=["always", "unavailable-only", "never"],
        help="fallback policy for auto engine selection",
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="show additional details during conversion",
    )
    parser.add_argument(
        "--progress",
        action="store_true",
        help="show a live progress bar on stderr (batch mode only)",
    )
    parser.add_argument(
        "--watch",
        action="store_true",
        help="re-convert when the source file changes (single-file mode only)",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    args = parser.parse_args(argv)

    options = ConversionOptions.from_environment()
    if args.fallback:
        options = replace(options, fallback=args.fallback)

    # ── Batch mode ────────────────────────────────────────────────────────────
    if args.output_dir is not None:
        if args.watch:
            parser.error("--watch is not supported in batch mode (--output-dir)")
        srcs = list(args.inputs)
        if not srcs:
            srcs = sorted(glob.glob("*.docx"))
            if not srcs:
                parser.error("no input files given and no .docx found in the current directory")
            if not args.quiet:
                print(
                    f"[docx2pdf-py] auto-discovered {len(srcs)} file(s)",
                    file=sys.stderr,
                )
        for src in srcs:
            if not os.path.exists(src):
                parser.error(f"file not found: {src}")

        pbar = _ProgressBar(len(srcs)) if args.progress else None

        def _progress(item) -> None:  # type: ignore[no-untyped-def]
            if pbar is not None:
                pbar.advance()
            if args.quiet:
                return
            if item.succeeded:
                extra = ""
                if args.verbose and item.result:
                    r = item.result
                    extra = (
                        f"  [engine: {r.engine} | "
                        f"{r.elapsed_seconds:.3f}s | "
                        f"pages: {r.page_count or '?'}]"
                    )
                print(f"OK  {item.input_path} -> {item.output_path}{extra}")
            elif item.cancelled:
                print(f"--  {item.input_path} (cancelled)", file=sys.stderr)
            else:
                print(f"ERR {item.input_path}: {item.error}", file=sys.stderr)

        results = convert_batch(
            srcs,
            args.output_dir,
            engine=args.engine,
            options=options,
            max_workers=args.workers,
            on_progress=_progress,
        )
        if pbar is not None:
            pbar.finish()
        failures = [r for r in results if r.failed]
        return 1 if failures else 0

    # ── Single-file mode ──────────────────────────────────────────────────────
    inputs = args.inputs
    if len(inputs) == 0:
        cands = sorted(glob.glob("*.docx"))
        if not cands:
            parser.error("no input file given and no .docx found in the current directory")
        src = cands[0]
        out = "output.pdf"
        if not args.quiet:
            print(f"[docx2pdf-py] auto-selected input: {src}", file=sys.stderr)
    elif len(inputs) == 1:
        src = inputs[0]
        out = "output.pdf"
    elif len(inputs) == 2:
        src, out = inputs[0], inputs[1]
    else:
        parser.error(
            "too many positional arguments for single-file mode; "
            "use --output-dir for batch conversion"
        )

    if not os.path.exists(src):
        parser.error(f"file not found: {src}")
    if os.path.exists(out) and not args.force and not args.watch:
        parser.error(f"output already exists: {out} (use -f to overwrite)")

    if args.verbose and not args.quiet:
        print(f"Input:  {src}")
        print(f"Output: {out}")
        print(f"Requested engine: {args.engine}")

    try:
        result = convert_detailed(src, out, engine=args.engine, options=options)
    except (Docx2PdfError, OSError) as exc:
        parser.error(str(exc))

    if not args.quiet:
        for warning in result.warnings:
            print(f"[docx2pdf-py] {warning}", file=sys.stderr)
        print(f"OK {src} -> {out}  [engine: {result.engine}]")
        if args.verbose:
            print(
                f"Elapsed: {result.elapsed_seconds:.3f}s | "
                f"Pages: {result.page_count or 'unknown'} | "
                f"Output: {result.output_bytes} bytes"
            )
            for attempt in result.attempts:
                status = attempt.error or ("available" if attempt.available else "unavailable")
                print(
                    f"Attempt: {attempt.engine} | {attempt.elapsed_seconds:.3f}s | {status}"
                )

    if args.watch:
        if not args.quiet:
            print(f"[docx2pdf-py] watching {src} for changes... (Ctrl+C to stop)", file=sys.stderr)
        last_mtime = os.path.getmtime(src)
        try:
            while True:
                time.sleep(1)
                try:
                    mtime = os.path.getmtime(src)
                except OSError:
                    continue
                if mtime != last_mtime:
                    last_mtime = mtime
                    if not args.quiet:
                        print("[docx2pdf-py] change detected, re-converting...", file=sys.stderr)
                    try:
                        convert_detailed(src, out, engine=args.engine, options=options)
                        if not args.quiet:
                            print(f"OK {src} -> {out}")
                    except (Docx2PdfError, OSError) as exc:
                        print(f"ERR {exc}", file=sys.stderr)
        except KeyboardInterrupt:
            if not args.quiet:
                print("\n[docx2pdf-py] stopped", file=sys.stderr)

    return 0


if __name__ == "__main__":
    sys.exit(main())
