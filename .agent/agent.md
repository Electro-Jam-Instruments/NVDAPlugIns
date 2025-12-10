# NVDA-Plugins Project - Agent Instructions

## Project Overview

This repository contains NVDA accessibility plugins, starting with **powerpoint-comments** - a plugin for accessible PowerPoint comment navigation.

**Repository:** `NVDA-Plugins` (multi-plugin structure)
**Current Plugin:** `powerpoint-comments/`

## Quick Context

- **Target:** PowerPoint 365 with Modern Comments only
- **User:** Screen reader user who cannot see the console
- **Tech Stack:** Python, comtypes (not pywin32), NVDA addon framework
- **COM for data, UIA for focus** - This is the core architectural decision

## Key Documents

| Document | Purpose |
|----------|---------|
| `MVP_IMPLEMENTATION_PLAN.md` | Current implementation plan (6 phases) |
| `REPO_STRUCTURE.md` | Multi-plugin repository structure |
| `.agent/experts/` | Domain knowledge and research |

## Expert Knowledge Areas

When working in specific domains, consult these expert files:

| Area | Expert File | Use For |
|------|-------------|---------|
| NVDA Plugin Development | `.agent/experts/nvda-plugins/nvda-plugins.md` | Addon structure, APIs, packaging |
| PowerPoint Automation | `.agent/experts/powerpoint-automation/powerpoint-automation.md` | COM API, comments, views |
| Windows Accessibility | `.agent/experts/windows-accessibility/windows-accessibility.md` | UIA, focus management |
| Local AI Vision | `.agent/experts/local-ai-vision/local-ai-vision.md` | Deferred - image descriptions |

Each expert folder contains:
- `{area}.md` - Distilled knowledge summary
- `decisions.md` - Key decisions and rationale
- `research/` - Original research documents

## Current MVP Phases

1. **Foundation + View Management** - App module, view detection, auto-switch to Normal
2. **Slide Change Detection** - Detect changes, announce comment status
3. **Focus First Comment** - UIA focus to Comments pane
4. **Comment Navigation** - Ctrl+Alt+PageUp/Down between comments
5. **@Mention Detection** - Find comments mentioning current user
6. **Polish + Packaging** - Error handling, .nvda-addon packaging

## Technical Decisions (Summary)

See individual `decisions.md` files in expert folders for full rationale.

| Decision | Choice | Why |
|----------|--------|-----|
| COM Library | comtypes | NVDA uses it internally; pywin32 has DLL issues |
| Comment Data | COM API | Reliable, works while file open |
| Focus Management | UIA | Comments pane is UIA-enabled |
| @Mention Parsing | Regex | No structured API available |
| View Detection | `ActiveWindow.ViewType` | Returns constants (Normal=9) |

## Accessibility Reminder

The user relies on a screen reader. When providing output:
- Keep responses concise
- Use clear structure
- Announce progress on multi-step tasks
- The frontend announces TODO status changes automatically

## File Organization

```
NVDA-Plugins/
├── .agent/                    # Agent knowledge base (this folder)
├── powerpoint-comments/       # First plugin (when we start coding)
├── test-resources/            # Test presentations
├── archive/                   # Old/superseded content (gitignored)
├── MVP_IMPLEMENTATION_PLAN.md
└── REPO_STRUCTURE.md
```

## Version

- **Last Updated:** December 2025
- **Status:** Planning complete, ready for Phase 1 implementation
