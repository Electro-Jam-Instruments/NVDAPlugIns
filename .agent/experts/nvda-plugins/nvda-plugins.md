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

**CRITICAL:** To extend NVDA's built-in PowerPoint support, use the EXACT NVDA documentation pattern:

```python
# appModules/powerpnt.py
# Pattern: NVDA Developer Guide - extending built-in appModules
# https://download.nvaccess.org/documentation/developerGuide.html

# Import EVERYTHING from built-in - this is the NVDA doc pattern
from nvdaBuiltin.appModules.powerpnt import *

# Inherit from just-imported AppModule (VERIFIED WORKING v0.0.9+)
class AppModule(AppModule):
    """Extended PowerPoint support."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)  # super() works for __init__
        # Your initialization here
```

**WARNING - PATTERNS THAT DO NOT WORK:**
```python
# WRONG - Explicit alias import (v0.0.4-v0.0.8) - Module does NOT load
from nvdaBuiltin.appModules.powerpnt import AppModule as BuiltinPowerPointAppModule
class AppModule(BuiltinPowerPointAppModule):  # Does not work!

# WRONG - Base class (v0.0.1-v0.0.3) - Loads but loses built-in features
from nvdaBuiltin.appModules.powerpnt import *
class AppModule(appModuleHandler.AppModule):  # Loses built-in support
```

Only the EXACT NVDA documentation pattern works. The explicit alias pattern appears logically equivalent but does not work in practice.

For full details on verified patterns, see `decisions.md` Decision 6.

### Event Handler Rules - CRITICAL

**`event_appModule_gainFocus` is an OPTIONAL HOOK - parent class does NOT define it.**

```python
def event_appModule_gainFocus(self):
    """Called when PowerPoint gains focus.

    CRITICAL RULES:
    1. Do NOT call super() - method doesn't exist in parent, will crash
    2. Do NOT do heavy work - blocks NVDA speech
    3. Defer COM/heavy work with core.callLater()
    """
    log.info("App gained focus - deferring initialization")
    core.callLater(100, self._deferred_initialization)

# WRONG - Will crash with AttributeError
def event_appModule_gainFocus(self):
    super().event_appModule_gainFocus()  # FAILS - method doesn't exist
```

**When to use super():**
- `__init__` - YES, parent has this
- `terminate` - YES, parent has this
- `event_appModule_gainFocus` - NO, optional hook
- `event_appModule_loseFocus` - NO, optional hook

**Why defer heavy work?**
Event handlers that block prevent NVDA from speaking. The 100ms delay with `core.callLater()` allows NVDA to complete focus handling and speak before our code runs.

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

### CRITICAL: Use comHelper, NOT direct GetActiveObject

**NVDA runs with UIAccess privileges** which prevents direct COM access to lower-privilege processes like PowerPoint. Direct `GetActiveObject()` fails with `WinError -2147221021 Operation unavailable`.

```python
# CORRECT - Use NVDA's comHelper (VERIFIED WORKING v0.0.13)
import comHelper
ppt_app = comHelper.getActiveObject("PowerPoint.Application", dynamic=True)

# WRONG - Fails with UIAccess privilege error
from comtypes.client import GetActiveObject
ppt_app = GetActiveObject("PowerPoint.Application")  # FAILS!
```

**Why comHelper works:**
- Uses in-process injection when available (bypasses privilege restrictions)
- Falls back to subprocess with appropriate privileges
- Same approach NVDA's built-in PowerPoint module uses

### Basic COM Pattern

```python
import comHelper

# Connect to running app (using comHelper!)
app = comHelper.getActiveObject("PowerPoint.Application", dynamic=True)

# Access COM properties
slide = app.ActiveWindow.View.Slide
comments = slide.Comments

# Iterate collections (1-indexed!)
for i in range(1, comments.Count + 1):
    comment = comments.Item(i)
    print(comment.Text)
```

## Threading for COM Operations (Recommended Pattern)

**NVDA maintainers recommend dedicated threads over `core.callLater()` for continuous or repeated work.**

### Why Use Threading?
- `core.callLater()` creates new deferred calls with no lifecycle management
- No cleanup when app closes or NVDA exits
- Continuous monitoring (like slide changes) needs a persistent thread

### Threading Pattern for NVDA Addons

```python
import threading
import queue
from comtypes import CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED
from queueHandler import queueFunction, eventQueue

class WorkerThread:
    def __init__(self):
        self._stop_event = threading.Event()
        self._work_queue = queue.Queue()
        self._thread = None

    def start(self):
        self._thread = threading.Thread(
            target=self._run,
            name="MyWorker",
            daemon=False  # Non-daemon for clean shutdown
        )
        self._thread.start()

    def stop(self, timeout=5):
        self._stop_event.set()
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=timeout)

    def queue_task(self, task_name, *args):
        self._work_queue.put((task_name, args))

    def _run(self):
        # CRITICAL: Initialize COM in STA mode for Office apps
        CoInitializeEx(COINIT_APARTMENTTHREADED)
        try:
            while not self._stop_event.is_set():
                try:
                    task_name, args = self._work_queue.get(timeout=0.5)
                    self._execute_task(task_name, args)
                except queue.Empty:
                    pass
        finally:
            CoUninitialize()  # Always cleanup COM

    def _announce(self, message):
        # CRITICAL: Thread-safe UI announcement
        queueFunction(eventQueue, ui.message, message)
```

### Threading Rules

| Rule | Reason |
|------|--------|
| `CoInitializeEx(COINIT_APARTMENTTHREADED)` | Office COM requires STA |
| `CoUninitialize()` in finally block | Prevents COM leaks |
| `threading.Event()` for stop signal | Clean shutdown |
| `daemon=False` | Allows cleanup before exit |
| `queueFunction(eventQueue, ...)` for UI | Thread-safe NVDA speech |
| Join with timeout in `terminate()` | Prevents hang on exit |

### AppModule Integration

```python
class AppModule(AppModule):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._worker = WorkerThread()
        self._worker.start()

    def event_appModule_gainFocus(self):
        self._worker.queue_task("initialize")

    def terminate(self):
        if self._worker:
            self._worker.stop(timeout=5)
        super().terminate()
```

For full details, see `decisions.md` Decision #12.

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
