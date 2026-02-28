# NVDA Addon Development Documentation

General knowledge for developing NVDA addons. For plugin-specific documentation, see the plugin's `docs/` folder.

## Contents

| Folder | Purpose |
|--------|---------|
| `guides/` | How-to guides for common tasks |
| `research/` | Deep technical research (general NVDA, not app-specific) |
| `experts/` | Expert agent definition files |
| `reference/` | API and pattern references |
| `history/` | Historical docs, completed plans |

## Experts

| Expert | Description |
|--------|-------------|
| [expert-uia.md](experts/expert-uia.md) | Windows UI Automation patterns for NVDA |

## Guides

| Guide | Description |
|-------|-------------|
| [nvda-development-guide.md](guides/nvda-development-guide.md) | Core NVDA addon development patterns |
| [nvda-testing-guide.md](guides/nvda-testing-guide.md) | Testing NVDA addons |

## Research

| Document | Description |
|----------|-------------|
| [NVDA_UIA_Deep_Research.md](research/NVDA_UIA_Deep_Research.md) | UIA integration in NVDA |

## Plugin-Specific Documentation

Each plugin has its own `docs/` folder with the same structure:
- `powerpoint-comments/docs/` - PowerPoint Comments plugin (comment navigation, presenter notes)
- `windows-dictation-silence/docs/` - Windows Dictation Silence plugin (Win+H voice typing without echo)

...and others as they are added.
