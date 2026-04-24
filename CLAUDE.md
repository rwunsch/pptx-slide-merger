# pptx-slide-merger

Pure-Python PPTX slide merger. Copies slides between PowerPoint files preserving original layouts, masters, themes, and media.

## Architecture

- `pptx_slide_merger/merger.py` — Core `PptxMerger` class plus standalone `reorder_slides()` and `move_slide()` functions. Works at ZIP/XML level using lxml.
- `pptx_slide_merger/cli.py` — CLI entry point (`pptx-merge` command) with subcommands: `merge`, `list`, `reorder`, `move`.
- `tests/test_merger.py` — Tests using pytest. Generates fixtures in-memory, no external PPTX files needed.

## Key OOXML Concepts

- PPTX files are ZIP archives containing XML parts.
- Slides reference layouts, layouts reference masters, masters reference themes.
- `sldMasterId` and `sldLayoutId` share a global ID namespace — must be unique across all masters.
- OOXML element ordering in `presentation.xml` is strict (sldMasterIdLst before sldIdLst before sldSz).
- Notes slides reference a notes master which must exist.
- Media relationships include image, video, audio, and the Microsoft `media` alt type.
- Slide order is determined by `<p:sldId>` element order in `sldIdLst` — reordering only requires rearranging those elements.

## Slide Reordering

- `reorder_slides(path, new_order, output_path=None)` — rearranges slides by rewriting `sldIdLst` in `presentation.xml`. Works directly on the ZIP, no temp extraction.
- `move_slide(path, from_index, to_index, output_path=None)` — convenience wrapper around `reorder_slides`.
- CLI: `pptx-merge reorder deck.pptx --order 2,0,1` and `pptx-merge move deck.pptx --from 2 --to 0`

## Commands

```bash
pip install -e .                  # install in dev mode
pip install -e ".[test]"          # install with test deps
pytest tests/                     # run tests
pptx-merge merge base.pptx -a source.pptx:0,1,2 -o output.pptx  # merge slides
pptx-merge list deck.pptx                                         # list slides
pptx-merge reorder deck.pptx --order 2,0,1 -o output.pptx        # reorder slides
pptx-merge move deck.pptx --from 2 --to 0 -o output.pptx         # move one slide
```
