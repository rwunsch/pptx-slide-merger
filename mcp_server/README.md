# pptx-slide-merger MCP Server

MCP server that exposes PowerPoint slide management tools to AI assistants.

## Status

**Not yet implemented.** See [architecture](../docs/architecture.md) for design decisions and planned tools.

## Planned Tools

| Tool | Description |
|------|-------------|
| `list_slides` | List slide titles with 0-based indices |
| `reorder_slides` | Rearrange all slides into a new order |
| `move_slide` | Move a single slide to a new position |
| `merge_slides` | Merge slides from multiple PPTX files |

## Installation (once implemented)

```bash
pip install -e ".[mcp]"
```

## Configuration (once implemented)

Add to Claude Code:

```bash
claude mcp add pptx-manager -s user -- python -m mcp_server
```

Or add to `~/.claude.json` manually:

```json
{
  "mcpServers": {
    "pptx-manager": {
      "command": "python",
      "args": ["-m", "mcp_server"],
      "cwd": "/path/to/pptx-slide-merger"
    }
  }
}
```
