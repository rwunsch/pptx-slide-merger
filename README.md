# pptx-slide-merger

Pure-Python PowerPoint slide management: merge slides across decks, reorder within a deck, and expose it all to AI assistants. No .NET runtime required.

## Components

| Component | Path | Description |
|-----------|------|-------------|
| **Core Library** | [`pptx_slide_merger/`](pptx_slide_merger/) | Python API + CLI for slide merging, reordering, and listing |
| **MCP Server** | [`mcp_server/`](mcp_server/) | MCP tools for AI assistants (planned) |
| **Skill** | [`skill/`](skill/) | Claude Code skill for CLI-based slide management |
| **Visual-edits diff** | [`scripts/diff_pptx_styling.py`](scripts/diff_pptx_styling.py) | Diff two PPTX files at shape/paragraph/run level — see [visual-edits workflow](docs/visual-edits-workflow.md) |

See [Architecture](docs/architecture.md) for design decisions and component relationships. Real-engagement patterns and learnings in [docs/learnings-from-real-engagements.md](docs/learnings-from-real-engagements.md).

## Quick Start

```bash
pip install -e .
```

### Python API

```python
from pathlib import Path
from pptx_slide_merger import PptxMerger, reorder_slides, move_slide, list_slides

# List slides
titles = list_slides(Path("deck.pptx"))

# Reorder slides
reorder_slides(Path("deck.pptx"), [2, 0, 1])

# Move a single slide
move_slide(Path("deck.pptx"), from_index=4, to_index=0)

# Merge from multiple decks
with PptxMerger(Path("base.pptx")) as merger:
    merger.add_slide(Path("source1.pptx"), 0)
    merger.add_slide(Path("source2.pptx"), 3)
    merger.save(Path("output.pptx"))
```

### CLI

```bash
pptx-merge list deck.pptx                                         # list slides
pptx-merge reorder deck.pptx --order 2,0,1 -o output.pptx        # reorder slides
pptx-merge move deck.pptx --from 2 --to 0 -o output.pptx         # move one slide
pptx-merge merge base.pptx -a source.pptx:0,1,2 -o output.pptx   # merge slides
```

## How It Works

PPTX files are ZIP archives containing XML parts. Each slide references a layout, which references a master, which references a theme.

- **Merging** operates at the ZIP/XML level: extracts, traces layout/master/theme chains, remaps global OOXML IDs, deduplicates media, and writes a valid PPTX
- **Reordering** is lightweight: rearranges `<p:sldId>` elements in `presentation.xml` within the ZIP. No extraction needed.

## Limitations

- Charts and OLE objects are not yet supported
- Slide animations/transitions may not be fully preserved
- Hyperlinks to other slides within the same deck are not remapped

## License

MIT
