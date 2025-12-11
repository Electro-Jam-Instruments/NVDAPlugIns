# NVDA Plugin Development - Decisions

## Decision Log

### 1. Use comtypes, not pywin32

**Decision:** Use `comtypes` for all COM automation
**Date:** December 2025
**Status:** Final

**Rationale:**
- NVDA uses comtypes internally
- pywin32 has DLL conflicts when loaded in NVDA process
- comtypes is already available in NVDA runtime

**Research:** `research/NVDA_PowerPoint_Native_Support_Analysis.md`

---

### 2. App Module Approach (not Global Plugin alone)

**Decision:** Create `appModules/powerpnt.py` as primary entry point
**Date:** December 2025
**Status:** Final

**Rationale:**
- Integrates with NVDA's existing PowerPoint support
- Can use overlay classes for PowerPoint-specific objects
- Inherits existing COM event infrastructure

**Alternatives Considered:**
- Global Plugin only: Would miss PowerPoint-specific events
- Hybrid: More complex, not needed for MVP

**Research:** `research/NVDA_PowerPoint_Native_Support_Analysis.md` Section 6

---

### 3. NVDA Version Compatibility

**Decision:** Target NVDA 2025.1+ with lastTested 2025.3.2
**Date:** December 2025
**Status:** Updated

**Rationale:**
- Always target latest stable NVDA release
- User's system runs NVDA 2025.3.2
- Ensures access to newest APIs and improvements
- NVDA 2025.1 adds IUIAutomation6, improved speech, Remote Access

**Version Policy:** Update minimum and lastTested versions when new stable NVDA releases become available.

---

### 4. No Legacy Comment Support

**Decision:** Only support Modern Comments (PowerPoint 365)
**Date:** December 2025
**Status:** Final

**Rationale:**
- Legacy comments use different COM API
- Modern comments are the future
- Simplifies implementation significantly
- Target users are on 365

---

### 5. manifest.ini Quoting Format

**Decision:** Use specific quoting rules for manifest.ini
**Date:** December 2025
**Status:** Final - LEARNED THE HARD WAY

**The Rules:**
- No quotes for single words: `name = addonName`
- Double quotes for text with spaces: `summary = "My Addon"`
- Triple quotes for multi-line: `description = """Long text"""`
- No quotes for versions/URLs: `version = 0.1.0`

**Why This Matters:**
- Incorrect quoting causes NVDA to silently reject the addon
- Error messages are not helpful
- Took significant debugging time to discover

**Common Failures:**
- `summary = My Addon Name` → FAILS (needs quotes)
- `name = "addonName"` → May work but incorrect
- Using smart quotes (""") instead of straight quotes (""") → FAILS

---

### 6. Extend Built-in PowerPoint Support (Don't Replace)

**Decision:** Use EXACT NVDA documentation pattern: `import *` then `class AppModule(AppModule)`
**Date:** December 2025
**Status:** VERIFIED WORKING v0.0.9+

**Rationale:**
- NVDA has ~1500 lines of existing PowerPoint support
- Replacing it would break working features
- Extending allows adding comments without losing existing functionality

**Version History - What We Learned:**
- v0.0.1-v0.0.3: Used `appModuleHandler.AppModule` - MODULE LOADED but lost built-in features
- v0.0.4-v0.0.8: Used explicit import with alias - MODULE DID NOT LOAD
- v0.0.9+: Uses EXACT NVDA doc pattern - WORKING

**PATTERNS THAT DON'T WORK (tested):**

```python
# PATTERN A: Inheriting from base class (v0.0.1-v0.0.3)
# Module LOADS but loses all built-in PowerPoint support!
from nvdaBuiltin.appModules.powerpnt import *
import appModuleHandler
class AppModule(appModuleHandler.AppModule):  # LOSES built-in features

# PATTERN B: Explicit import with alias (v0.0.4-v0.0.8)
# Module appears to NOT LOAD - no logs appear
from nvdaBuiltin.appModules.powerpnt import AppModule as BuiltinPowerPointAppModule
class AppModule(BuiltinPowerPointAppModule):  # Did not work in testing
```

**CORRECT PATTERN (NVDA Developer Guide):**

```python
# EXACT NVDA DOCUMENTATION PATTERN (v0.0.9+)
# Reference: https://download.nvaccess.org/documentation/developerGuide.html

from nvdaBuiltin.appModules.powerpnt import *

class AppModule(AppModule):  # Inherits from the just-imported AppModule!
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)  # super() works for __init__
```

**Why This Specific Pattern:**
1. `import *` brings the built-in's `AppModule` into our namespace
2. `class AppModule(AppModule):` creates our class inheriting from the imported one
3. Our new class SHADOWS the imported `AppModule` name
4. NVDA finds our `AppModule` and uses it
5. We inherit all built-in functionality

**References:**
- NVDA Developer Guide wwahost example: The ONLY documented pattern
- Joseph Lee's Office Desk addon: Uses just `import *` with no class (re-exports built-in)

**Why This Matters:**
- Pattern B (explicit alias) did not work despite appearing correct
- Only the EXACT doc pattern has been verified to work
- Module loading is silent - no errors shown when pattern is wrong
- Took versions v0.0.1 through v0.0.9 to debug properly

---

### 7. Logging Strategy for Event Debugging

**Decision:** Use Python logging module extensively in Phase 1
**Date:** December 2025
**Status:** Final

**Rationale:**
- Events may not fire as expected
- Screen reader users cannot see console output
- NVDA log provides persistent debugging record
- Can verify behavior without visual feedback

**Implementation:**
```python
import logging
log = logging.getLogger(__name__)

log.debug("Event fired")
log.info("Important action")
log.error(f"Failed: {e}")
```

**View logs:** NVDA menu > Tools > View Log (NVDA+F1)

---

### 8. Testing Strategy - Manual First, Automated Later

**Decision:** Use manual NVDA testing for MVP, consider automation post-MVP
**Date:** December 2025
**Status:** Final

**Rationale:**
- Automated NVDA testing tools exist but are complex to set up
- Manual testing with scratchpad is fastest for iteration
- Real screen reader testing catches issues automation misses
- Automation useful for regression testing after MVP stable

**Manual Testing Workflow:**
1. Copy to scratchpad: `%APPDATA%\nvda\scratchpad\appModules\`
2. Enable scratchpad in NVDA settings
3. Reload plugins: NVDA+Ctrl+F3
4. Check NVDA log for errors

**Post-MVP Automation Options:**
- NVDA Testing Driver (C#)
- Guidepup (JavaScript)

---

## Backlogged Decisions

### Comment Resolution Status (Deferred)

**Issue:** Cannot reliably detect resolved vs unresolved comments
**Status:** Backlogged

**Why Deferred:**
- Resolved status not exposed in COM API
- Would require OOXML parsing (file is locked while open)
- Shadow copy approach is complex and fragile

**Research:** `research/PowerPoint-Comment-Resolution-LockedFile-Access-Research.md`

**Future Option:** Revisit if Microsoft exposes resolution status in COM API

---

### 9. Event Handler super() Rules

**Decision:** Only call super() on methods that exist in parent class
**Date:** December 2025
**Status:** VERIFIED v0.0.10-v0.0.11

**The Problem:**
- v0.0.10 added `super().event_appModule_gainFocus()` assuming it would preserve parent behavior
- Crashed with `AttributeError: 'super' object has no attribute 'event_appModule_gainFocus'`
- The parent class does NOT define this method - it's an optional hook

**The Rule:**

| Method | Call super()? | Reason |
|--------|---------------|--------|
| `__init__` | YES | Parent has this |
| `terminate` | YES | Parent has this |
| `event_appModule_gainFocus` | NO | Optional hook, parent doesn't have it |
| `event_appModule_loseFocus` | NO | Optional hook, parent doesn't have it |

**CORRECT:**
```python
def __init__(self, *args, **kwargs):
    super().__init__(*args, **kwargs)  # YES - parent has __init__

def event_appModule_gainFocus(self):
    # NO super() call - method doesn't exist in parent
    core.callLater(100, self._deferred_work)
```

**WRONG:**
```python
def event_appModule_gainFocus(self):
    super().event_appModule_gainFocus()  # CRASH - AttributeError
```

---

### 10. Defer Heavy Work in Event Handlers

**Decision:** Use `core.callLater()` to defer COM work from event handlers
**Date:** December 2025
**Status:** VERIFIED WORKING v0.0.11+

**The Problem:**
- v0.0.9 blocked NVDA speech by doing COM work directly in `event_appModule_gainFocus`
- Event handlers that block prevent NVDA from speaking focus changes

**The Solution:**
```python
def event_appModule_gainFocus(self):
    # Return immediately - don't block NVDA speech
    core.callLater(100, self._deferred_initialization)

def _deferred_initialization(self):
    # COM work happens here, 100ms after focus event completes
    self._ppt_app = comHelper.getActiveObject("PowerPoint.Application", dynamic=True)
```

**Why 100ms?**
- Not from official NVDA guidance - it's a pragmatic value
- Allows NVDA to complete focus handling and speak first
- Could be 50ms or 200ms - the key is ANY deferral

**Note on core.callLater():**
- NVDA maintainers recommend dedicated threads over `core.callLater()` for continuous work
- For one-time initialization like ours, `core.callLater()` is acceptable
- Future: Consider dedicated thread for continuous comment monitoring

---

### 11. Use comHelper for COM Access (UIAccess Privilege)

**Decision:** Use `comHelper.getActiveObject()` NOT direct `GetActiveObject()`
**Date:** December 2025
**Status:** VERIFIED WORKING v0.0.13

**The Problem:**
- v0.0.11-v0.0.12 failed with `WinError -2147221021 Operation unavailable`
- NVDA runs with UIAccess privileges
- Windows prevents high-privilege processes from directly accessing COM objects in lower-privilege processes

**The Solution:**
```python
# CORRECT
import comHelper
ppt_app = comHelper.getActiveObject("PowerPoint.Application", dynamic=True)

# WRONG - Fails with UIAccess error
from comtypes.client import GetActiveObject
ppt_app = GetActiveObject("PowerPoint.Application")
```

**Why comHelper works:**
- Uses in-process injection when available (NVDAHelper's `nvdaInProcUtils_getActiveObject`)
- Falls back to subprocess with appropriate privileges
- Same approach NVDA's built-in PowerPoint module uses

**Reference:** NVDA GitHub Issue #2483 - GetActiveObject fails when running with uiAccess

---

### 12. Use Dedicated Background Thread for COM Operations

**Decision:** Move all COM operations to a dedicated background thread instead of using `core.callLater()`
**Date:** December 2025
**Status:** VERIFIED WORKING v0.0.14

**The Problem:**
- `core.callLater()` is not recommended by NVDA maintainers for continuous/repeated work
- Each focus event creates a new deferred call with no lifecycle management
- No cleanup when PowerPoint closes or NVDA exits
- Phase 2 needs continuous slide monitoring - `callLater` won't scale

**The Solution:**

```python
import threading
import queue
from comtypes import CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED
from queueHandler import queueFunction, eventQueue

class PowerPointWorker:
    """Background thread for COM operations."""

    def __init__(self):
        self._stop_event = threading.Event()
        self._work_queue = queue.Queue()
        self._thread = None
        self._ppt_app = None

    def start(self):
        self._thread = threading.Thread(
            target=self._run,
            name="PowerPointCommentWorker",
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
        # Initialize COM in STA mode (required for Office)
        CoInitializeEx(COINIT_APARTMENTTHREADED)
        try:
            while not self._stop_event.is_set():
                try:
                    task_name, args = self._work_queue.get(timeout=0.5)
                    self._execute_task(task_name, args)
                except queue.Empty:
                    pass
        finally:
            self._ppt_app = None
            CoUninitialize()

    def _announce(self, message):
        # Thread-safe UI announcement
        queueFunction(eventQueue, ui.message, message)
```

**Threading Rules for NVDA Addons:**

| Rule | Reason |
|------|--------|
| Use `CoInitializeEx(COINIT_APARTMENTTHREADED)` | Office COM requires STA |
| Always `CoUninitialize()` in finally block | Prevents COM leaks |
| Use `threading.Event()` for stop signal | Clean shutdown |
| Non-daemon thread (`daemon=False`) | Allows cleanup before exit |
| Use `queueFunction(eventQueue, ...)` for UI | Thread-safe NVDA speech |
| Join with timeout in `terminate()` | Prevents hang on exit |

**Why This Pattern:**
1. COM initialized once per thread (proper STA)
2. Work queue allows task dispatch from main thread
3. Thread can run continuously for Phase 2 monitoring
4. Clean shutdown via `terminate()` method
5. Thread-safe announcements via `queueHandler`

**AppModule Integration:**

```python
class AppModule(AppModule):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._worker = PowerPointWorker()
        self._worker.start()

    def event_appModule_gainFocus(self):
        # Queue task instead of core.callLater
        self._worker.queue_task("initialize")

    def terminate(self):
        # Clean shutdown
        if self._worker:
            self._worker.stop(timeout=5)
        super().terminate()
```

**Key Insight:** NVDA maintainers recommend dedicated threads over `core.callLater()` for anything beyond simple one-shot operations. This pattern is reusable for future NVDA plugins.

---

### 13. Define EApplication Interface Locally (NOT from Type Library)

**Decision:** Define the PowerPoint EApplication COM events interface locally, do NOT use type library loading
**Date:** December 2025
**Status:** VERIFIED WORKING v0.0.21 - COM events firing successfully

**The Problem:**
- v0.0.16-v0.0.20 tried to load PowerPoint type library to get event interfaces
- ALL approaches failed with `[WinError -2147319779] Library not registered`
- Tried: GetModule by app object, GUID lookup, registry path - all failed

**The Discovery:**
- NVDA's own `nvdaBuiltin/appModules/powerpnt.py` defines `EApplication` class LOCALLY
- It does NOT import from a type library
- The class defines interface GUID and DISPIDs manually
- `wireEApplication` does NOT exist - this was a misunderstanding of NVDA's code

**The Solution - Define Interface Locally:**

```python
import comtypes
from comtypes import IDispatch, COMObject
from comtypes.client._events import _AdviseConnection
import ctypes

class EApplication(IDispatch):
    """PowerPoint Application events interface - defined locally."""
    _iid_ = comtypes.GUID("{914934C2-5A91-11CF-8700-00AA0060263B}")  # Events interface GUID
    _methods_ = []
    _disp_methods_ = [
        comtypes.DISPMETHOD([comtypes.dispid(2001)], None, "WindowSelectionChange",
            (["in"], ctypes.POINTER(IDispatch), "sel")),
        comtypes.DISPMETHOD([comtypes.dispid(2013)], None, "SlideShowNextSlide",
            (["in"], ctypes.POINTER(IDispatch), "slideShowWindow")),
    ]

class PowerPointEventSink(COMObject):
    _com_interfaces_ = [EApplication, IDispatch]

    def WindowSelectionChange(self, sel):
        # Called when slide selection changes in edit mode
        pass

    def SlideShowNextSlide(self, slideShowWindow):
        # Called during slideshow navigation
        pass

# Connection (on STA thread with message pump):
sink = PowerPointEventSink()
sink_iunknown = sink.QueryInterface(comtypes.IUnknown)
connection = _AdviseConnection(ppt_app, EApplication, sink_iunknown)
```

**Critical GUIDs and DISPIDs:**

| Item | Value | Notes |
|------|-------|-------|
| EApplication Interface GUID | `{914934C2-5A91-11CF-8700-00AA0060263B}` | USE THIS |
| Type Library GUID | `{91493440-5A91-11CF-8700-00AA0060263B}` | DO NOT USE - causes "Library not registered" |
| WindowSelectionChange DISPID | 2001 | Edit mode slide changes |
| SlideShowNextSlide DISPID | 2013 | Slideshow navigation |

**Why Type Library Loading Fails:**
- PowerPoint's type library may not be registered properly on all systems
- Office 365 deployment may not register type libraries in expected locations
- GUID-based loading is less reliable than app-object-based loading
- Even when type library loads, wire interfaces have naming conflicts

**What NVDA Does (source: powerpnt.py):**
1. Defines `EApplication(IDispatch)` class locally with GUID and DISPIDs
2. Creates `PowerPointEventSink(COMObject)` implementing the interface
3. Uses `_AdviseConnection()` to connect sink to PowerPoint Application
4. Runs Windows message pump to receive events

**What We Tried That Failed:**

| Version | Approach | Result |
|---------|----------|--------|
| v0.0.16 | `GetModule(ppt_app)` | `Library not registered` |
| v0.0.17 | Fixed bad import | Same error |
| v0.0.18 | Multiple fallbacks (GUID, registry) | All failed |
| v0.0.19 | Access `wireEApplication` via `import *` | Not exported |
| v0.0.20 | Access via `hasattr(module, 'wireEApplication')` | Not found |

**Research Document:** `.agent/experts/nvda-plugins/research/PowerPoint-COM-Events-Research.md`

**Key Insight:** When COM type library loading fails, define the interface locally. This is exactly what NVDA does, and it works reliably because it doesn't depend on system type library registration.

---
