# Architecture

## Overview

pptx-slide-merger is organized as three components sharing one repository:

```
pptx-slide-merger/
  pptx_slide_merger/       # Core Python library
  mcp_server/              # MCP server (wraps the library)
  skill/                   # Claude Code skill (wraps the CLI)
  tests/                   # All tests
```

## Components

### Core Library (`pptx_slide_merger/`)

The foundation. Pure-Python PPTX manipulation at the ZIP/XML level using lxml. No .NET runtime required.

**Responsibilities:**
- Cross-deck slide merging (`PptxMerger` class) — copies slides preserving layouts, masters, themes, and media
- Within-deck slide reordering (`reorder_slides`, `move_slide`) — lightweight functions that rewrite `sldIdLst` element order in `presentation.xml`
- Slide listing (`list_slides`) — reads slide titles via python-pptx
- CLI (`pptx-merge`) — subcommands: `merge`, `list`, `reorder`, `move`

**Key design decisions:**
- Operates at ZIP/XML level rather than using python-pptx for mutations, because python-pptx doesn't support cross-deck operations or low-level relationship management
- python-pptx is used only for reading (listing slides, test verification)
- Reorder/move functions work directly on the ZIP without temp directory extraction, making them fast and safe for in-place edits

### MCP Server (`mcp_server/`)

Exposes the core library as MCP tools for use by AI assistants (Claude Code, Cursor, Windsurf, custom agents).

**Key design decisions:**
- **Framework: fastmcp** — chosen over the lower-level `mcp` SDK because the tool surface is small and well-defined. fastmcp's decorator-based approach keeps the server code minimal. The lower-level SDK would add boilerplate without benefit for this use case.
- **Imports the core library directly** — no subprocess/CLI wrapping. This gives typed parameters, proper error propagation, and avoids shell escaping issues with file paths.
- **Stateless per-call** — each tool invocation is independent. No session state or file handles held between calls.

**Planned tools:**
- `list_slides(file)` — list slide titles with indices
- `reorder_slides(file, order, output?)` — full reorder
- `move_slide(file, from, to, output?)` — single slide move
- `merge_slides(base, sources, output)` — cross-deck merge

### Skill (`skill/`)

A Claude Code skill (markdown instructions) that guides Claude to use the `pptx-merge` CLI for slide management tasks.

**Key design decisions:**
- **Wraps the CLI, not the library** — skills are instruction files, not code. Claude executes CLI commands via Bash.
- **Complements the MCP server** — when the MCP server is configured, Claude uses it directly. When it's not available (e.g., quick terminal session), the skill provides a fallback via CLI.
- **Includes workflow guidance** — not just "how to call the tool" but "list slides first, confirm with user, then reorder." This makes Claude's behavior more predictable and user-friendly.

## Dependency Graph

```
skill/ ──uses──> CLI (pptx-merge)
                      │
                      ▼
mcp_server/ ──imports──> pptx_slide_merger/
                              │
                              ├── merger.py (lxml, zipfile)
                              └── cli.py (argparse)
```

## Packaging

One `pyproject.toml` at the root:
- `pip install .` — core library + CLI
- `pip install ".[mcp]"` — adds fastmcp dependency for the MCP server
- `pip install ".[test]"` — adds pytest for testing
- The skill requires no installation — it's a markdown file copied or symlinked into `~/.claude/skills/`
