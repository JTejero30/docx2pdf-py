"""Command-line interface: docx2pdf-py [inputs...] [output] [--output-dir DIR]"""
from __future__ import annotations

import argparse
import glob
import os
import sys
from collections.abc import Sequence
from dataclasses import replace

from . import __version__
from .api import convert_batch, convert_detailed
from .exceptions import Docx2PdfError
from .models import ConversionOptions


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
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    args = parser.parse_args(argv)

    options = ConversionOptions.from_environment()
    if args.fallback:
        options = replace(options, fallback=args.fallback)

    # ── Batch mode ────────────────────────────────────────────────────────────
    if args.output_dir is not None:
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

        def _progress(item) -> None:  # type: ignore[no-untyped-def]
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
    if os.path.exists(out) and not args.force:
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

    return 0


if __name__ == "__main__":
    sys.exit(main())
