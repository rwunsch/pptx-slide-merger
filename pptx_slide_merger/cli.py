#!/usr/bin/env python3
"""Command-line interface for pptx-slide-merger."""

import argparse
import sys
from pathlib import Path

from .merger import PptxMerger, list_slides, reorder_slides, move_slide


def _cmd_merge(args):
    """Original merge behavior."""
    if not args.add:
        print("Error: no slides specified. Use -a file.pptx:0,1,2", file=sys.stderr)
        sys.exit(1)

    with PptxMerger(args.base, verbose=not args.quiet) as merger:
        for spec in args.add:
            if ':' in spec:
                src_path, indices_str = spec.rsplit(':', 1)
                src = Path(src_path)
                indices = [int(i.strip()) for i in indices_str.split(',')]
            else:
                src = Path(spec)
                indices = list(range(len(list_slides(src))))

            if not src.exists():
                print(f"Warning: {src} not found, skipping", file=sys.stderr)
                continue

            for idx in indices:
                ok = merger.add_slide(src, idx)
                if not args.quiet:
                    status = "+" if ok else "!"
                    print(f"  {status} {src.name} slide {idx}")

        merger.save(args.output)


def _cmd_list(args):
    """List slides in a PPTX file."""
    titles = list_slides(args.file)
    for i, t in enumerate(titles):
        print(f"  [{i:3d}] {t}")
    print(f"\n{len(titles)} slides")


def _cmd_reorder(args):
    """Reorder slides in a PPTX file."""
    indices = [int(i.strip()) for i in args.order.split(',')]
    output = args.output or args.file
    reorder_slides(args.file, indices, output)
    if not args.quiet:
        print(f"Reordered {args.file.name} -> {output.name}")


def _cmd_move(args):
    """Move a single slide to a new position."""
    output = args.output or args.file
    move_slide(args.file, args.from_index, args.to_index, output)
    if not args.quiet:
        print(f"Moved slide {args.from_index} -> position {args.to_index} "
              f"in {output.name}")


def main():
    parser = argparse.ArgumentParser(
        description="PPTX slide merger and manager.",
    )
    subparsers = parser.add_subparsers(dest="command")

    # --- merge ---
    merge_p = subparsers.add_parser(
        "merge",
        help="Merge slides from multiple PPTX files into one",
        epilog="Example: pptx-merge merge base.pptx -a source1.pptx:0,1,2 -o merged.pptx",
    )
    merge_p.add_argument("base", type=Path, help="Base PPTX file")
    merge_p.add_argument("-a", "--add", action="append", default=[],
                         help="Source file and slide indices: file.pptx:0,1,2")
    merge_p.add_argument("-o", "--output", type=Path, default=Path("merged.pptx"),
                         help="Output file (default: merged.pptx)")
    merge_p.add_argument("-q", "--quiet", action="store_true")

    # --- list ---
    list_p = subparsers.add_parser("list", help="List slides in a PPTX file")
    list_p.add_argument("file", type=Path, help="PPTX file to list")

    # --- reorder ---
    reorder_p = subparsers.add_parser(
        "reorder", help="Reorder slides in a PPTX file")
    reorder_p.add_argument("file", type=Path, help="PPTX file to reorder")
    reorder_p.add_argument("--order", required=True,
                           help="Comma-separated slide indices in new order: 2,0,1")
    reorder_p.add_argument("-o", "--output", type=Path, default=None,
                           help="Output file (default: overwrite input)")
    reorder_p.add_argument("-q", "--quiet", action="store_true")

    # --- move ---
    move_p = subparsers.add_parser(
        "move", help="Move a single slide to a new position")
    move_p.add_argument("file", type=Path, help="PPTX file")
    move_p.add_argument("--from", dest="from_index", type=int, required=True,
                        help="Current slide index (0-based)")
    move_p.add_argument("--to", dest="to_index", type=int, required=True,
                        help="Target position (0-based)")
    move_p.add_argument("-o", "--output", type=Path, default=None,
                        help="Output file (default: overwrite input)")
    move_p.add_argument("-q", "--quiet", action="store_true")

    args = parser.parse_args()

    if args.command is None:
        parser.print_help()
        sys.exit(1)

    if args.command == "merge":
        if not args.base.exists():
            print(f"Error: {args.base} not found", file=sys.stderr)
            sys.exit(1)
        _cmd_merge(args)
    elif args.command == "list":
        if not args.file.exists():
            print(f"Error: {args.file} not found", file=sys.stderr)
            sys.exit(1)
        _cmd_list(args)
    elif args.command == "reorder":
        if not args.file.exists():
            print(f"Error: {args.file} not found", file=sys.stderr)
            sys.exit(1)
        _cmd_reorder(args)
    elif args.command == "move":
        if not args.file.exists():
            print(f"Error: {args.file} not found", file=sys.stderr)
            sys.exit(1)
        _cmd_move(args)


if __name__ == "__main__":
    main()
