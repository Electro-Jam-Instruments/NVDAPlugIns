# NVDA Plugin Development - Expert Knowledge

## Overview

This document contains distilled knowledge for developing NVDA addons, specifically for PowerPoint integration.

## Target NVDA Version

| Field | Value |
|-------|-------|
| Minimum Version | 2025.1 |
| Last Tested | 2025.3.2 |
| Documentation | https://www.nvaccess.org/files/nvda/documentation/developerGuide.html |

**Policy:** Always target the latest stable NVDA release. Update these values when new stable versions are released.

**Note:** User's system runs NVDA 2025.3.2, comtypes 1.4.11, Python 3.11.9 (32-bit).

## Reference Files

| Topic | File | What It Contains |
|-------|------|------------------|
| Decisions | `decisions.md` | Rationale for technical choices |
| Current Implementation | `MVP_IMPLEMENTATION_PLAN.md` | Phase-by-phase code |
| Directory Structure | `REPO_STRUCTURE.md` | Addon folder layout, manifest template |
| Build & Package | `nvda-addon-packager` agent | Build scripts, version management |
| Testing | `accessibility-tester` agent | Test strategies, debugging |
| Research | `research/` folder | Deep analysis documents |

## NVDA Addon Structure

```
addon/
├── manifest.ini          # Required: addon metadata
├── appModules/           # App-specific modules (by exe name)
│   └── powerpnt.py       # PowerPoint module
├── globalPlugins/        # Always-active plugins
│   └── myPlugin/
│       └── __init__.py
├── doc/                  # Documentation
│   └── en/
│       └── readme.html
└── locale/               # Translations
    └── en/
        └── LC_MESSAGES/
```

## Key Concepts

### App Modules vs Global Plugins

| Type | When Active | Use For |
|------|-------------|---------|
| App Module | Only when app has focus | App-specific features |
| Global Plugin | Always | Cross-app features |

**For PowerPoint:** Use App Module (`appModules/powerpnt.py`)

### CRITICAL: manifest.ini Quoting Rules

**This is a common source of errors that can take hours to debug!**

| Field Type | Quote Style | Example |
|------------|-------------|---------|
| Single word (no spaces) | No quotes | `name = addonName` |
| Single line WITH spaces | `"double quotes"` | `summary = "My Addon Name"` |
| Multi-line text | `"""triple quotes"""` | `description = """Long text here"""` |
| Version numbers | No quotes | `version = 0.1.0` |
| URLs | No quotes | `url = https://github.com/...` |

**Common Mistakes:**
- Using quotes everywhere - WRONG for single words
- Forgetting quotes on text with spaces - WILL FAIL
- Using smart quotes instead of straight quotes - WILL FAIL

**If NVDA rejects your addon, check the manifest quoting first!**

For full manifest template, see `REPO_STRUCTURE.md`.

## NVDA APIs

### Core Imports

```python
import appModuleHandler      # Base class for app modules
from scriptHandler import script  # Keyboard shortcut decorator
import ui                    # Speech output
import tones                 # Audio feedback
import api                   # Current focus, navigator object
from NVDAObjects import NVDAObject  # Base object class
import controlTypes          # Control type constants
```

### Speech Output

```python
import ui

ui.message("Text to speak")           # Speak immediately
ui.message("Text", speechPriority=1)  # High priority
```

### Keyboard Scripts

```python
from scriptHandler import script

@script(
    description="Description for input help",
    gesture="kb:control+alt+c",
    category="PowerPoint Comments"
)
def script_myAction(self, gesture):
    # Implementation
    pass
```

### Audio Feedback

```python
import tones

tones.beep(440, 100)   # Frequency Hz, duration ms
tones.beep(880, 50)    # Higher pitch, shorter
```

### Focus and Navigation

```python
import api

obj = api.getFocusObject()        # Currently focused object
nav = api.getNavigatorObject()    # Navigator object
api.setFocusObject(obj)           # Move focus
```

## PowerPoint-Specific

### NVDA's PowerPoint Architecture

NVDA's built-in `powerpnt.py` (~1500 lines) provides:
- COM event handling via `EApplication` sink
- Overlay classes: `Slide`, `Shape`, `TextFrame`, `Table`
- `TextFrameTextInfo` for text navigation
- Slideshow mode handling

**Key limitation:** No native comment support

### Extending Built-in App Modules

**IMPORTANT:** To extend NVDA's built-in PowerPoint support (not replace it):

```python
# appModules/powerpnt.py
from nvdaBuiltin.appModules.powerpnt import *  # Inherit all existing support
```

This pattern inherits all existing PowerPoint support so we only add comment features on top.

For full implementation, see `MVP_IMPLEMENTATION_PLAN.md` Phase 1.

### Logging for Debugging

Always add logging to track event firing:

```python
import logging
log = logging.getLogger(__name__)

log.debug("Method called")
log.info("Important event")
log.error("Something failed")
```

View logs: NVDA menu > Tools > View Log (or NVDA+F1)

### Overlay Classes

```python
def chooseNVDAObjectOverlayClasses(self, obj, clsList):
    """Add custom classes to NVDA objects."""
    if self._is_comment(obj):
        clsList.insert(0, CommentObject)
```

## COM Integration (comtypes)

### Why comtypes, not pywin32

- NVDA uses comtypes internally
- pywin32 DLLs conflict with NVDA process
- comtypes already in NVDA runtime

### Basic Pattern

```python
from comtypes.client import GetActiveObject, CreateObject

# Connect to running app
app = GetActiveObject("PowerPoint.Application")

# Access COM properties
slide = app.ActiveWindow.View.Slide
comments = slide.Comments

# Iterate collections (1-indexed!)
for i in range(1, comments.Count + 1):
    comment = comments.Item(i)
    print(comment.Text)
```

## Testing During Development

### Scratchpad Testing (Fastest Iteration)

1. Copy module to: `%APPDATA%\nvda\scratchpad\appModules\powerpnt.py`
2. Enable in NVDA: Settings > Advanced > Developer Scratchpad
3. Reload plugins: NVDA+Ctrl+F3
4. Check NVDA log for errors (NVDA+F1)

For full testing workflow, see the `accessibility-tester` agent.

## Common Gotchas

1. **COM collections are 1-indexed** - Use `range(1, count + 1)`
2. **COM calls can fail** - Always wrap in try/except
3. **Don't block the main thread** - Use threading for long operations
4. **Test with actual screen reader users** - Keyboard-only testing misses issues

## Research Files

See `research/` folder for detailed analysis:
- `NVDA_PowerPoint_Native_Support_Analysis.md` - How NVDA handles PowerPoint
- `NVDA-PowerPoint-Community-Addons-Research.md` - Existing addons
- `NVDA_UIA_Deep_Research.md` - UIA integration details
