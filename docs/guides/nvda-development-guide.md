# NVDA Addon Development Guide

General knowledge for developing NVDA addons. For app-specific details, see the plugin's docs folder.

## Target NVDA Version

| Field | Value |
|-------|-------|
| Minimum Version | 2025.1 |
| Last Tested | 2025.3.2 |
| Documentation | https://www.nvaccess.org/files/nvda/documentation/developerGuide.html |

**Policy:** Always target the latest stable NVDA release.

## NVDA Addon Structure

```
addon/
├── manifest.ini          # Required: addon metadata
├── appModules/           # App-specific modules (by exe name)
│   └── appname.py        # Named after executable
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

## App Modules vs Global Plugins

| Type | When Active | Use For |
|------|-------------|---------|
| App Module | Only when app has focus | App-specific features |
| Global Plugin | Always | Cross-app features |

## manifest.ini Quoting Rules

**This is a common source of errors!**

| Field Type | Quote Style | Example |
|------------|-------------|---------|
| Single word (no spaces) | No quotes | `name = addonName` |
| Single line WITH spaces | `"double quotes"` | `summary = "My Addon Name"` |
| Multi-line text | `"""triple quotes"""` | `description = """Long text here"""` |
| Version numbers | No quotes | `version = 0.1.0` |
| URLs | No quotes | `url = https://github.com/...` |

**If NVDA rejects your addon, check the manifest quoting first!**

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
    category="My Addon"
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

## Event Handler Rules

### Optional Hooks

Some event handlers are optional hooks - the parent class does NOT define them:

```python
def event_appModule_gainFocus(self):
    """Called when app gains focus.

    CRITICAL RULES:
    1. Do NOT call super() - method doesn't exist in parent, will crash
    2. Do NOT do heavy work - blocks NVDA speech
    3. Defer heavy work with worker thread or core.callLater()
    """
    pass

# WRONG - Will crash with AttributeError
def event_appModule_gainFocus(self):
    super().event_appModule_gainFocus()  # FAILS - method doesn't exist
```

**When to use super():**
- `__init__` - YES, parent has this
- `terminate` - YES, parent has this
- `event_appModule_gainFocus` - NO, optional hook
- `event_appModule_loseFocus` - NO, optional hook

## NVDA Event Sequence

### Event Order (When Object Gains Focus)

| Order | Event/Method | Purpose |
|-------|--------------|---------|
| 1 | `chooseNVDAObjectOverlayClasses(obj, clsList)` | Select overlay classes |
| 2 | `event_NVDAObject_init(obj)` | Modify object properties |
| 3 | `event_gainFocus(obj, nextHandler)` | Handle focus event |

### event_NVDAObject_init

Fires BEFORE NVDA announces the object. Modify `obj.name` here and NVDA speaks the modified name.

```python
def event_NVDAObject_init(self, obj):
    """Modify object properties BEFORE announcement.

    Available only in App Modules (not Global Plugins).
    """
    if should_modify(obj):
        obj.name = f"custom prefix, {obj.name}"
```

### Overlay Classes

```python
def chooseNVDAObjectOverlayClasses(self, obj, clsList):
    """Add custom classes to NVDA objects."""
    if self._is_my_object(obj):
        clsList.insert(0, MyCustomObject)
```

## Logging for Debugging

```python
import logging
log = logging.getLogger(__name__)

log.debug("Method called")
log.info("Important event")
log.error("Something failed")
```

View logs: NVDA menu > Tools > View Log (or NVDA+F1)

## COM Integration (comtypes)

### Why comtypes, not pywin32

- NVDA uses comtypes internally
- pywin32 DLLs conflict with NVDA process
- comtypes already in NVDA runtime

### CRITICAL: Use comHelper for COM Access

NVDA runs with UIAccess privileges which prevents direct COM access.

```python
# CORRECT - Use NVDA's comHelper
import comHelper
app = comHelper.getActiveObject("Application.Name", dynamic=True)

# WRONG - Fails with UIAccess privilege error
from comtypes.client import GetActiveObject
app = GetActiveObject("Application.Name")  # FAILS!
```

### COM Collections are 1-indexed

```python
# Iterate collections (1-indexed!)
for i in range(1, collection.Count + 1):
    item = collection.Item(i)
```

## Threading for COM Operations

NVDA maintainers recommend dedicated threads for continuous work.

### Threading Pattern

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

    def _run(self):
        # CRITICAL: Initialize COM in STA mode for Office apps
        CoInitializeEx(COINIT_APARTMENTTHREADED)
        try:
            while not self._stop_event.is_set():
                try:
                    task = self._work_queue.get(timeout=0.5)
                    self._execute_task(task)
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

## Testing During Development

### Scratchpad Testing (Optional - Fast Iteration)

1. Copy module to: `%APPDATA%\nvda\scratchpad\appModules\`
2. Enable in NVDA: Settings > Advanced > Developer Scratchpad
3. Reload plugins: NVDA+Ctrl+F3
4. Check NVDA log for errors (NVDA+F1)

### Full Addon Testing

1. Build .nvda-addon package (scons)
2. Install via double-click
3. Restart NVDA

## Common Gotchas

1. **COM collections are 1-indexed** - Use `range(1, count + 1)`
2. **COM calls can fail** - Always wrap in try/except
3. **Don't block the main thread** - Use threading for long operations
4. **Test with actual screen reader** - Keyboard-only testing misses speech issues
5. **Check manifest quoting** - Most common addon install failure
