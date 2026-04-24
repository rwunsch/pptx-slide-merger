---
name: pptx-manager
description: Manage PowerPoint slides — list, reorder, move, and merge slides using the pptx-merge CLI
---

# PPTX Slide Manager

Use the `pptx-merge` CLI to manage PowerPoint slides. The tool is installed as part of the pptx-slide-merger package.

## Available Commands

```bash
pptx-merge list <file.pptx>                                    # List all slides with indices
pptx-merge reorder <file.pptx> --order 2,0,1 [-o output.pptx]  # Reorder slides
pptx-merge move <file.pptx> --from 3 --to 0 [-o output.pptx]   # Move one slide
pptx-merge merge <base.pptx> -a <src.pptx>:0,1,2 -o out.pptx   # Merge from multiple decks
```

## Workflow

1. **Always list slides first** before any reorder/move operation, so you and the user can see the current order with indices.
2. **Confirm with the user** before modifying files, especially for in-place operations (no `-o` flag).
3. **Use `-o` for a new file** when the user wants to keep the original unchanged.
4. **Omit `-o`** to modify the file in-place (the default for `reorder` and `move`).

## Indices

All slide indices are **0-based**. When showing slides to the user, display them with their index so they can reference them easily.

## Examples

**User:** "What slides are in my deck?"
```bash
pptx-merge list presentation.pptx
```

**User:** "Move the last slide to the front"
```bash
# First list to find the index, then move
pptx-merge list presentation.pptx
pptx-merge move presentation.pptx --from 9 --to 0
```

**User:** "Reverse the slide order"
```bash
# For a 5-slide deck: reverse is 4,3,2,1,0
pptx-merge reorder presentation.pptx --order 4,3,2,1,0
```

**User:** "Combine slides from two decks"
```bash
pptx-merge merge base.pptx -a other.pptx:0,2,4 -o combined.pptx
```
