from __future__ import annotations

import argparse
import json
import sys
import threading
from pathlib import Path

from .engine import ConversionCancelled, ConversionEngine
from .models import ConversionOptions


def parse_toc_level(value: str) -> int | None:
    if value.lower() == "auto":
        return None
    try:
        level = int(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("toc level must be 'auto' or a positive integer") from exc
    if level < 1:
        raise argparse.ArgumentTypeError("toc level must be positive")
    return level


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="markforge", description="Convert local documents into structured Markdown.")
    commands = parser.add_subparsers(dest="command", required=True)
    convert = commands.add_parser("convert", help="Convert a file or directory")
    convert.add_argument("input", type=Path)
    convert.add_argument("--output", type=Path, required=True)
    convert.add_argument("--split", choices=("chapters", "pages", "single"), default="chapters")
    convert.add_argument("--processing-pages", type=int, default=20)
    convert.add_argument("--output-pages", type=int, default=25)
    convert.add_argument("--toc-level", type=parse_toc_level, default=None, metavar="auto|N")
    convert.add_argument("--device", choices=("auto", "mps", "cuda", "cpu"), default="auto")
    convert.add_argument("--combined", action="store_true")
    convert.add_argument("--force", action="store_true")
    convert.add_argument("--extractor", choices=("marker", "native"), default="marker")

    inspect_command = commands.add_parser("inspect", help="Inspect a PDF without converting it")
    inspect_command.add_argument("input", type=Path)
    inspect_command.add_argument("--split", choices=("chapters", "pages", "single"), default="chapters")
    inspect_command.add_argument("--output-pages", type=int, default=25)
    inspect_command.add_argument("--toc-level", type=parse_toc_level, default=None, metavar="auto|N")
    inspect_command.add_argument("--device", choices=("auto", "mps", "cuda", "cpu"), default="auto")
    commands.add_parser("gui", help="Launch the desktop application")
    return parser


def _progress(event: dict[str, object]) -> None:
    if event.get("type") == "segment_started":
        print(f"[{event.get('index')}/{event.get('total')}] {event.get('title')} (pages {event.get('start_page')}-{event.get('end_page')})")
    elif event.get("type") == "segment_completed":
        print(f"  {event.get('status')}: {event.get('title')}")
    elif event.get("type") == "document_completed":
        print(f"Manifest: {event.get('manifest')}")


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.command == "gui":
        from .gui import main as gui_main
        gui_main()
        return 0
    try:
        if args.command == "inspect":
            options = ConversionOptions(
                output_dir=Path("out"), split=args.split, output_pages=args.output_pages,
                toc_level=args.toc_level, device=args.device,
            )
            print(json.dumps(ConversionEngine().inspect(args.input, options).to_dict(), indent=2, ensure_ascii=False))
            return 0
        options = ConversionOptions(
            output_dir=args.output, split=args.split, processing_pages=args.processing_pages,
            output_pages=args.output_pages, toc_level=args.toc_level, device=args.device,
            combined=args.combined, force=args.force, extractor=args.extractor,
        )
        manifests = ConversionEngine(progress=_progress, cancel_event=threading.Event()).convert(args.input, options)
        print(f"Converted {len(manifests)} document(s).")
        failed = 0
        for manifest_path in manifests:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            failed += sum(
                segment.get("status") in {"failed", "cancelled"}
                for segment in manifest.get("segments", [])
            )
        if failed:
            print(
                f"{failed} segment(s) failed or were cancelled. See manifest.json for details.",
                file=sys.stderr,
            )
            return 2
        return 0
    except ConversionCancelled:
        print("Conversion cancelled.", file=sys.stderr)
        return 130
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
