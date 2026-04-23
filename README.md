# pptx-slide-merger

Pure-Python library for copying slides between PowerPoint files while preserving original layouts, masters, themes, and media. No .NET runtime required.

## Installation

```bash
pip install -e .
```

## Usage

### Python API

```python
from pathlib import Path
from pptx_slide_merger import PptxMerger

with PptxMerger(Path("base.pptx")) as merger:
    merger.add_slide(Path("source1.pptx"), 0)   # slide index 0
    merger.add_slide(Path("source2.pptx"), 3)   # slide index 3
    merger.add_slide(Path("base.pptx"), 5)      # from base deck too
    merger.save(Path("output.pptx"))
```

List slide titles:

```python
from pptx_slide_merger import list_slides
titles = list_slides(Path("deck.pptx"))
```

### CLI

```bash
# List slides in a file
pptx-merge deck.pptx --list

# Merge specific slides (0-based indices)
pptx-merge base.pptx -a source1.pptx:0,1,2 -a source2.pptx:3,4 -o merged.pptx

# Merge all slides from a source
pptx-merge base.pptx -a source.pptx -o merged.pptx
```

## How It Works

PPTX files are ZIP archives containing XML parts. Each slide references a layout, which references a master, which references a theme. The merger operates at the ZIP/XML level:

1. Extracts the base PPTX, strips its slides but keeps masters/layouts/themes
2. For each added slide, traces its layout -> master -> theme chain
3. Copies the entire master with all its layouts when first encountered
4. Remaps global OOXML IDs (sldMasterId, sldLayoutId) to avoid collisions
5. Deduplicates media files by content hash
6. Handles images, video, audio, speaker notes, and notes masters
7. Strips comment references (comments are not copied)
8. Writes a valid PPTX with correct Content_Types and element ordering

## Limitations

- Charts and OLE objects are not yet supported
- Slide animations/transitions may not be fully preserved
- Hyperlinks to other slides within the same deck are not remapped
- Very large decks (500+ slides) may be slow due to per-file XML parsing
