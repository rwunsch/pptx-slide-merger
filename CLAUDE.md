# pptx-slide-merger

Pure-Python PPTX slide merger. Copies slides between PowerPoint files preserving original layouts, masters, themes, and media.

## Architecture

- `pptx_slide_merger/merger.py` — Core `PptxMerger` class. Works at ZIP/XML level using lxml.
- `pptx_slide_merger/cli.py` — CLI entry point (`pptx-merge` command).
- `tests/test_merger.py` — Tests using pytest. Generates fixtures in-memory, no external PPTX files needed.

## Key OOXML Concepts

- PPTX files are ZIP archives containing XML parts.
- Slides reference layouts, layouts reference masters, masters reference themes.
- `sldMasterId` and `sldLayoutId` share a global ID namespace — must be unique across all masters.
- OOXML element ordering in `presentation.xml` is strict (sldMasterIdLst before sldIdLst before sldSz).
- Notes slides reference a notes master which must exist.
- Media relationships include image, video, audio, and the Microsoft `media` alt type.

## Commands

```bash
pip install -e .                  # install in dev mode
pip install -e ".[test]"          # install with test deps
pytest tests/                     # run tests
pptx-merge base.pptx -a source.pptx:0,1,2 -o output.pptx  # CLI usage
```
