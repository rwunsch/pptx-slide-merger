#!/usr/bin/env python3
"""Command-line interface for pptx-slide-merger."""

import argparse
import sys
from pathlib import Path

from .merger import PptxMerger, list_slides


def main():
    parser = argparse.ArgumentParser(
        description="Merge slides from multiple PPTX files into one.",
        epilog="Example: pptx-merge base.pptx -a source1.pptx:0,1,2 -a source2.pptx:3,4 -o merged.pptx",
    )
    parser.add_argument("base", type=Path, help="Base PPTX file (provides dimensions and default master)")
    parser.add_argument("-a", "--add", action="append", default=[],
                        help="Source file and slide indices: file.pptx:0,1,2 (0-based)")
    parser.add_argument("-o", "--output", type=Path, default=Path("merged.pptx"),
                        help="Output PPTX file (default: merged.pptx)")
    parser.add_argument("-l", "--list", action="store_true",
                        help="List slides in the base file and exit")
    parser.add_argument("-q", "--quiet", action="store_true",
                        help="Suppress progress output")

    args = parser.parse_args()

    if not args.base.exists():
        print(f"Error: {args.base} not found", file=sys.stderr)
        sys.exit(1)

    if args.list:
        titles = list_slides(args.base)
        for i, t in enumerate(titles):
            print(f"  [{i:3d}] {t}")
        print(f"\n{len(titles)} slides")
        sys.exit(0)

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


if __name__ == "__main__":
    main()
