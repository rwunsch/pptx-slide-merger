# Roadmap: pptx-slide-merger -> PowerPoint MCP

## Phase 1: Slide Merger (Current - v0.1)

- [x] Copy slides between PPTX files preserving original layouts/masters/themes
- [x] Handle media files (images, EMF, SVG)
- [x] Preserve speaker notes with notes master
- [x] Globally unique OOXML IDs (sldMasterId, sldLayoutId)
- [x] Correct OOXML element ordering
- [x] CLI tool (`pptx-merge`)
- [ ] Automated tests with sample PPTX files
- [ ] Media deduplication (avoid copying identical images twice)
- [ ] Chart/OLE object support

## Phase 2: Slide Inspection & Analysis (v0.2)

- [ ] Extract text content from all slides
- [ ] List shapes, images, and their positions
- [ ] Report slide layout and master usage
- [ ] Detect duplicate slides across decks
- [ ] Export slide thumbnails

## Phase 3: Slide Modification (v0.3)

- [ ] Replace text in existing slides (find/replace)
- [ ] Update images in placeholder shapes
- [ ] Reorder slides within a deck
- [ ] Delete specific slides
- [ ] Clone slides within the same deck

## Phase 4: Template-based Generation (v0.4)

- [ ] Create slides from templates with data binding
- [ ] Table population from structured data
- [ ] Chart creation from data
- [ ] Batch generation from CSV/JSON

## Phase 5: MCP Server (v1.0)

Transform into a Model Context Protocol server that enables AI assistants
to fully understand and manipulate PowerPoint files.

### MCP Tools

- `list_presentations` - List available PPTX files
- `inspect_slide` - Get full details of a specific slide
- `merge_slides` - Copy slides between presentations
- `modify_slide` - Update text, images, shapes
- `create_slide` - Generate a slide from a template
- `export_pdf` - Convert to PDF (via LibreOffice)
- `validate_pptx` - Check OOXML validity

### MCP Resources

- `pptx://{path}/slides` - Browse slides as resources
- `pptx://{path}/masters` - Browse masters and layouts
- `pptx://{path}/media` - Browse embedded media

### Architecture

```
AI Assistant (Claude, etc.)
    |
    v
MCP Protocol (stdio/SSE)
    |
    v
pptx-mcp-server
    |
    +-- pptx_slide_merger (this library)
    +-- python-pptx (read/write)
    +-- lxml (XML manipulation)
    +-- LibreOffice (PDF export, optional)
```

## Design Principles

1. **Pure Python** - No .NET, no Java, no external services
2. **Preserve fidelity** - Never degrade source formatting
3. **OOXML-correct** - Valid files that open without repair
4. **Composable** - Each capability works independently
5. **MCP-native** - Designed for AI assistant integration
