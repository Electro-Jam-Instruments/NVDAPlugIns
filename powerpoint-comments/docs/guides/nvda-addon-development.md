# NVDA Addon Development Guide

How to build and modify NVDA addons, specifically for PowerPoint integration.

## Addon Structure

```
addon/
├── manifest.ini          # Required: addon metadata
├── appModules/           # App-specific modules (by exe name)
│   └── powerpnt.py       # PowerPoint module
├── globalPlugins/        # Always-active plugins (not used here)
└── doc/                  # Documentation (optional)
```

## manifest.ini Format

**Quoting rules are critical - incorrect quoting causes silent failures!**

| Field Type | Quote Style | Example |
|------------|-------------|---------|
| Single word (no spaces) | No quotes | `name = addonName` |
| Text WITH spaces | `"double quotes"` | `summary = "My Addon Name"` |
| Multi-line text | `"""triple quotes"""` | `description = """Long text"""` |
| Version numbers | No quotes | `version = 0.1.0` |
| URLs | No quotes | `url = https://github.com/...` |

## The Inheritance Pattern - CRITICAL

**USE THE EXACT NVDA DOCUMENTATION PATTERN:**

```python
# CORRECT - Verified working v0.0.9+
from nvdaBuiltin.appModules.powerpnt import *

class AppModule(AppModule):  # Inherits from just-imported AppModule
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
```

**PATTERNS THAT DO NOT WORK:**

```python
# WRONG - Explicit alias (v0.0.4-v0.0.8) - Module does NOT load
from nvdaBuiltin.appModules.powerpnt import AppModule as BuiltinAppModule
class AppModule(BuiltinAppModule):
    pass

# WRONG - Base class (v0.0.1-v0.0.3) - Loads but loses built-in features
from nvdaBuiltin.appModules.powerpnt import *
class AppModule(appModuleHandler.AppModule):
    pass
```

## Event Handler Rules

**`event_appModule_gainFocus` is an OPTIONAL HOOK - parent class does NOT define it.**

```python
# CORRECT - No super() call, delegate to worker thread
def event_appModule_gainFocus(self):
    # Do NOT call super() - method doesn't exist in parent
    # Do NOT do heavy work - blocks NVDA speech
    # Delegate to worker thread (non-blocking)
    if self._worker:
        self._worker.request_initialize()

# WRONG - Will crash with AttributeError
def event_appModule_gainFocus(self):
    super().event_appModule_gainFocus()  # FAILS
```

**When to use super():**

| Method | Call super()? | Reason |
|--------|---------------|--------|
| `__init__` | YES | Parent has this |
| `terminate` | YES | Parent has this |
| `event_appModule_gainFocus` | NO | Optional hook |
| `event_appModule_loseFocus` | NO | Optional hook |

## Core NVDA APIs

```python
import ui                    # Speech output
import api                   # Current focus, navigator object
import speech                # Cancel queued speech
from scriptHandler import script  # Keyboard shortcuts
from queueHandler import queueFunction, eventQueue  # Thread-safe UI

# Speak to user
ui.message("Text to speak")

# Get focused object
obj = api.getFocusObject()

# Thread-safe announcement (from worker thread)
queueFunction(eventQueue, ui.message, "Text")
```

## Keyboard Scripts

```python
from scriptHandler import script

@script(
    description="Read slide notes",
    gesture="kb:control+alt+n",
    category="PowerPoint Comments"
)
def script_readNotes(self, gesture):
    # Implementation
    pass
```

## Overlay Classes

Modify NVDA's representation of objects:

```python
def chooseNVDAObjectOverlayClasses(self, obj, clsList):
    """Add custom classes to NVDA objects."""
    if self._is_slide(obj):
        clsList.insert(0, CustomSlide)

class CustomSlide(NVDAObject):
    def _get_name(self):
        # Lazy evaluation - called when NVDA needs the name
        return f"has notes, {self.name}"
```

## Threading for COM Operations

**Use dedicated worker threads for COM operations - this is the recommended pattern.**

> **Note:** Early versions used `core.callLater(100, ...)` to defer work. This is deprecated in favor of worker threads which provide better lifecycle management and no arbitrary delays.

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
            name="PowerPointWorker",
            daemon=False  # Non-daemon for clean shutdown
        )
        self._thread.start()

    def stop(self, timeout=5):
        self._stop_event.set()
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=timeout)

    def _run(self):
        CoInitializeEx(COINIT_APARTMENTTHREADED)  # STA for Office
        try:
            while not self._stop_event.is_set():
                try:
                    task = self._work_queue.get(timeout=0.5)
                    self._execute_task(task)
                except queue.Empty:
                    pass
        finally:
            CoUninitialize()

    def _announce(self, message):
        queueFunction(eventQueue, ui.message, message)
```

## Building the Addon

1. Update version in `manifest.ini` and `buildVars.py`
2. Run `scons` to build `.nvda-addon` file
3. Double-click to install, restart NVDA

## Testing During Development

**Scratchpad method (fastest iteration):**

1. Copy module to: `%APPDATA%\nvda\scratchpad\appModules\powerpnt.py`
2. Enable in NVDA: Settings > Advanced > Developer Scratchpad
3. Reload plugins: NVDA+Ctrl+F3
4. Check NVDA log: NVDA+F1

## Logging

```python
import logging
log = logging.getLogger(__name__)

log.debug("Verbose info")
log.info("Important event")
log.error(f"Error: {e}")
```

View logs: NVDA menu > Tools > View Log

## How to Know if Module Loaded

**Success indicators in NVDA log:**
```
INFO - appModules.powerpnt: PowerPoint addon loading (v0.1.0)
INFO - appModules.powerpnt: Built-in powerpnt imported successfully
INFO - appModules.powerpnt: PowerPoint Comments AppModule initializing
```

**Failure indicators:**
- No log entries from your addon = module didn't load (check inheritance pattern)
- `ImportError` or `ModuleNotFoundError` = bad import
- `AttributeError: 'super' object has no attribute...` = bad super() call

## Complete Example

For a full working skeleton with all pieces together, see [Complete Skeleton Guide](complete-skeleton.md).
