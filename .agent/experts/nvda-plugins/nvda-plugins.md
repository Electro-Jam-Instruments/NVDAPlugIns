# NVDA Plugin Development - Expert Knowledge

## Overview

This document contains distilled knowledge for developing NVDA addons, specifically for PowerPoint integration.

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

### manifest.ini Template

```ini
name = addon-name
summary = Short description
description = Longer description of features
author = Your Name
version = 1.0.0
url = https://github.com/...
minimumNVDAVersion = 2023.1
lastTestedNVDAVersion = 2024.4
```

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

### Extending PowerPoint Support

```python
# appModules/powerpnt.py
import appModuleHandler
from comtypes.client import GetActiveObject

class AppModule(appModuleHandler.AppModule):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._ppt_app = None

    def event_appModule_gainFocus(self):
        self._connect()

    def _connect(self):
        try:
            self._ppt_app = GetActiveObject("PowerPoint.Application")
        except Exception:
            self._ppt_app = None
```

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

## Packaging

### Build Process

1. Zip the `addon/` directory contents
2. Rename `.zip` to `.nvda-addon`
3. Users double-click to install

### Testing During Development

- NVDA Scratchpad: `%APPDATA%\nvda\scratchpad\`
- Enable in NVDA settings: Advanced > Developer Scratchpad
- Copy `appModules/powerpnt.py` to scratchpad
- Restart NVDA or reload plugins (NVDA+Ctrl+F3)

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
