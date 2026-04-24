# Roadmap

## v0.1 (done)

- [x] Copy slides between PPTX files preserving layouts, masters, and themes
- [x] Handle images, video, audio, and other media
- [x] Media deduplication by content hash
- [x] Preserve speaker notes with notes master
- [x] Globally unique OOXML IDs (sldMasterId, sldLayoutId)
- [x] Correct OOXML element ordering in presentation.xml
- [x] Strip comment references from copied slides
- [x] CLI tool (`pptx-merge`)

## v0.2 (done)

- [x] Slide reordering within a deck (`reorder_slides`, `move_slide`)
- [x] CLI subcommands (`merge`, `list`, `reorder`, `move`)
- [x] Claude Code skill for CLI-based slide management
- [x] Architecture documentation

## v0.3 (next)

- [ ] MCP server with fastmcp (list, reorder, move, merge tools)
- [ ] Improved error reporting for corrupt source files

## Suggestions

The following are ideas for future development, not committed plans:

- **Delete slides** — remove specific slides from a deck by index. Same lightweight approach as reorder (rewrite `sldIdLst` minus the deleted entries, then remove the orphaned slide XML/rels from the ZIP).
- **Duplicate/clone slides** — copy a slide within the same deck. Useful for template-based workflows where you stamp out variations of a base slide.
- **Text search/replace** — find and replace text across all slides. Useful for template variable substitution (e.g., `{{client_name}}` -> `Acme Corp`).
- **Chart support** — copy chart parts and their embedded data when merging slides. Currently charts are silently dropped.
- **OLE object support** — preserve embedded Excel/Word objects during merge.
- **Slide metadata/tags** — read and write custom metadata on slides. Would enable workflows like "move all slides tagged 'appendix' to the end."
- **Batch operations via JSON** — accept a JSON manifest describing multiple operations (reorder, merge, delete) to apply atomically. Useful for MCP-driven workflows that build up a complex edit.
- **PyPI publishing** — publish to PyPI as `pptx-slide-merger` for easy installation.
