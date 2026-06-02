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
pptx-merge review deck.pptx -o review/ --serve                   # build + serve review viewer
```

## Reviewer Mode (visual deck review)

`pptx-merge review` builds a self-contained, browser-based review viewer: it renders
each slide to a PNG, extracts per-shape geometry from the OOXML, and emits a static
HTML viewer with a commenting overlay. Reviewers click a slide to drop a comment that
is **slide-aware** (index + title), **location-aware** (x/y % of the slide), and
**shape-aware** (the nearest shape's name, text, and bounding box) — so a follow-up
CLI/agent pass can apply each change to the right shape on the right slide.

```bash
pptx-merge review deck.pptx                  # build <deck>-review/ (prints how to serve)
pptx-merge review deck.pptx -o review/ --serve          # build AND serve with auto-save
pptx-merge review deck.pptx -o review/ --serve --port 8200
```

**Review toggle (opt-in, collapsed by default).** The viewer opens as a clean slide
browser. Click the **🖉 Review** toggle in the header (or press **`R`**) to enter
Reviewer Mode: comment pins appear, the panel actions show, and clicking a slide drops
a comment. Toggle it off (or press `R` / ✕) to return to clean viewing — no pins, no
overlay.

**Auto-save (no manual Export needed).** An open comment bubble commits automatically
**2 seconds after the last keystroke**, and on **click-outside** (the outside click is
consumed, so it never drops a stray second comment). Each save writes to `localStorage`
and best-effort `PUT`s `review-comments.json` to the server.

**Folder persistence via the save-server.** `python3 -m http.server` is static and
can't accept the reviewer's writes. The `--serve` flag (and the generated
`serve-review.py` shim) start a small save-server that serves the viewer dir **and**
handles `PUT/POST /review-comments.json`, writing it straight into the viewer folder.
So auto-save lands on disk with no manual step. On startup the viewer auto-loads any
`review-comments.json` sitting next to it.

```bash
# Serve a previously-built viewer dir with the save-server (auto-save -> disk):
python3 review/serve-review.py --port 8000          # the dropped-in shim, no install needed
python3 -m pptx_slide_merger.review_server review/ --port 8000   # or via the package

# Static fallback (Export / Copy-for-Claude only, no auto-save to disk):
python3 -m http.server -d review/ 8000
```

**Keyboard shield.** Typing in a comment field never triggers viewer navigation or
shortcuts — letters like `r`/`x` and the arrow keys type normally into the comment
instead of toggling review mode or changing slides.

**Export / Copy-for-Claude.** The panel keeps **⤓ Export** (download
`review-comments.json`) and **⧉ Copy for Claude** (markdown to clipboard) as fallbacks
for when the viewer is served statically.

Add `review-comments.json` to the deck folder's `.gitignore` — it's reviewer working
state, not a deck source artifact.

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
