from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .converter import convert_folder


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="docs_to_markdown",
        description="Batch convert .doc/.docx files in a folder to Markdown.",
    )
    parser.add_argument(
        "input_dir",
        type=Path,
        help="Folder containing .doc/.docx files.",
    )
    parser.add_argument(
        "output_dir",
        nargs="?",
        type=Path,
        default=None,
        help='Output folder (default: "<input_dir>/markdown").',
    )
    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        help="Search subfolders recursively.",
    )
    parser.add_argument(
        "--include-doc",
        action="store_true",
        help="Also attempt to convert legacy .doc files (requires LibreOffice).",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite existing .md files.",
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="Launch GUI for file/folder selection.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)

    # Launch GUI if --gui flag is set
    if args.gui:
        from .gui import main as gui_main
        return gui_main(argv)

    output_dir = args.output_dir or (args.input_dir / "markdown")

    report = convert_folder(
        input_dir=args.input_dir,
        output_dir=output_dir,
        recursive=args.recursive,
        include_doc=args.include_doc,
        overwrite=args.overwrite,
    )

    print(f"Converted: {report.converted}")
    print(f"Skipped:   {report.skipped}")
    print(f"Failed:    {report.failed}")

    if report.failures:
        print("\nFailures:")
        for item in report.failures:
            print(f"- {item}")

    return 0 if report.failed == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
