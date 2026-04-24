# pptx-slide-merger Skill

Claude Code skill for managing PowerPoint slides via the `pptx-merge` CLI.

## Installation

Copy or symlink the skill file into your Claude Code skills directory:

```bash
# Symlink (recommended — stays in sync with repo)
ln -s "$(pwd)/skill/pptx-manager.md" ~/.claude/skills/pptx-manager.md
```

## What It Does

When you ask Claude Code to work with PowerPoint slides, this skill guides it to:

1. List slides first so you can see what's in the deck
2. Confirm the operation with you before modifying files
3. Use the appropriate `pptx-merge` subcommand (`reorder`, `move`, `merge`, `list`)

## Prerequisites

The core library must be installed:

```bash
pip install -e /path/to/pptx-slide-merger
```

## Usage

Once installed, Claude Code will automatically use the skill when you mention slide management tasks:

- "List the slides in presentation.pptx"
- "Move slide 5 to the beginning"
- "Reorder the slides to put the summary first"
- "Merge slides from two decks"
