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
| NVDA Addon Packager | `.agent/experts/nvda-addon-packager/nvda-addon-packager.md` | Build .nvda-addon files, validate manifests |
| Local AI Vision | `.agent/experts/local-ai-vision/local-ai-vision.md` | Deferred - image descriptions |
| Accessibility Tester | `.agent/experts/accessibility-tester/accessibility-tester.md` | Test strategies, debugging, log verification |

Each expert folder contains:
- `{area}.md` - Distilled knowledge summary
- `decisions.md` - Key decisions and rationale
- `research/` - Original research documents

## Current MVP Phases

1. **Foundation + View Management** - App module, view detection, auto-switch to Normal
1.1. **Package + Deploy Pipeline** - Build script, GitHub release, install on test system
2. **Slide Change Detection** - Detect changes, announce comment status
3. **Focus First Comment** - UIA focus to Comments pane
3.1. **Slide Navigation from Comments** - Navigate slides while in Comments pane
4. **@Mention Detection** - Find comments mentioning current user
5. **Polish + Packaging** - Error handling, final release
6. **Comment Navigation (optional)** - If arrow keys prove insufficient

## CRITICAL: AppModule Inheritance Pattern

**USE ONLY THE EXACT NVDA DOCUMENTATION PATTERN:**

```python
from nvdaBuiltin.appModules.powerpnt import *

class AppModule(AppModule):  # Inherits from just-imported AppModule
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)  # super() works for __init__
```

**DO NOT USE THESE - THEY DO NOT WORK:**
- `from ... import AppModule as Alias` then `class AppModule(Alias):` - Module does NOT load
- `class AppModule(appModuleHandler.AppModule):` - Loads but loses built-in features

This was verified through v0.0.1-v0.0.9 testing. See `decisions.md` Decision #6 for full history.

## CRITICAL: Event Handler Rules

**`event_appModule_gainFocus` is an OPTIONAL HOOK - parent does NOT define it:**

```python
def event_appModule_gainFocus(self):
    # Do NOT call super() - method doesn't exist, will crash
    # Do NOT do heavy work - blocks NVDA speech
    core.callLater(100, self._deferred_initialization)
```

See `decisions.md` Decisions #9 and #10.

## CRITICAL: COM Access Pattern

**Use `comHelper.getActiveObject()` NOT direct `GetActiveObject()`:**

```python
import comHelper
ppt_app = comHelper.getActiveObject("PowerPoint.Application", dynamic=True)
```

Direct `GetActiveObject()` fails with UIAccess privilege error. See `decisions.md` Decision #11.

## Technical Decisions (Summary)

See individual `decisions.md` files in expert folders for full rationale.

| Decision | Choice | Why |
|----------|--------|-----|
| COM Library | comtypes | NVDA uses it internally; pywin32 has DLL issues |
| Comment Data | COM API | Reliable, works while file open |
| Focus Management | UIA | Comments pane is UIA-enabled |
| @Mention Parsing | Regex | No structured API available |
| View Detection | `ActiveWindow.ViewType` | Returns constants (Normal=9) |
| **Extend Built-in** | **`import *` then `class AppModule(AppModule):`** | **ONLY pattern that works - see decisions.md #6** |
| Testing Strategy | Manual first, automation post-MVP | Fastest iteration, real SR testing |
| Debugging | Python logging to NVDA log | Verify events without visual feedback |

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
- **Current Version:** v0.0.16
- **Status:** Phase 2 Slide Change Detection via COM Events
  - AppModule loads using NVDA doc pattern
  - COM connection working via comHelper
  - Dedicated background thread for COM operations (v0.0.14)
  - Thread-safe UI announcements via queueHandler
  - Clean shutdown with terminate() method
  - Presentation detection and view switching working
  - **Phase 2 (v0.0.16):** COM events for slide detection (replaces polling)
  - **Phase 2 (v0.0.16):** SlideSelectionChanged + WindowSelectionChange events
  - **Phase 2 (v0.0.16):** Windows message pump for event delivery
  - **Phase 2 (v0.0.16):** Comprehensive error logging
  - Comment counting on slide change
  - Auto-open Comments pane when comments exist
  - Ready for testing COM event-driven detection
