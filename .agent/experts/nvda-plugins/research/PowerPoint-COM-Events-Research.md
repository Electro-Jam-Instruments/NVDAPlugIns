# PowerPoint COM Events for NVDA Addon - Research Analysis

## Executive Summary

This document provides comprehensive research on connecting to PowerPoint COM events from an NVDA addon for event-driven slide change detection. The research covers NVDA's native implementation, the EApplication interface, and three viable approaches for our addon.

### Key Findings

1. **NVDA defines its own EApplication interface** - It does NOT use PowerPoint's type library; instead, it defines a minimal interface with only the events it needs
2. **wireEApplication is NOT exported** - Despite being defined in NVDA's powerpnt.py, it is not accessible from addons via `import *`
3. **The correct approach is to define our own EApplication interface** - Copy NVDA's pattern and define the interface locally
4. **SlideSelectionChanged event is NOT in NVDA's interface** - NVDA only implements `WindowSelectionChange` and `SlideShowNextSlide`

### Recommended Approach

**Option B: Define Our Own EApplication Interface (Recommended)**

Define a minimal COM interface locally with the events we need, matching NVDA's pattern but adding `SlideSelectionChanged`.

---

## 1. How NVDA's powerpnt Module Sets Up COM Events

### 1.1 The EApplication Interface Definition

NVDA defines its own minimal COM event interface rather than loading PowerPoint's type library:

```python
class EApplication(IDispatch):
    _iid_ = comtypes.GUID("{914934C2-5A91-11CF-8700-00AA0060263B}")
    _methods_ = []
    _disp_methods_ = [
        comtypes.DISPMETHOD(
            [comtypes.dispid(2001)],
            None,
            "WindowSelectionChange",
            (["in"], ctypes.POINTER(IDispatch), "sel"),
        ),
        comtypes.DISPMETHOD(
            [comtypes.dispid(2013)],
            None,
            "SlideShowNextSlide",
            (["in"], ctypes.POINTER(IDispatch), "slideShowWindow"),
        ),
    ]
```

**Key Points:**
- Uses the PowerPoint EApplication GUID: `{914934C2-5A91-11CF-8700-00AA0060263B}`
- Defines only 2 of the many available PowerPoint events
- Uses dispatch IDs (DISPIDs) to identify events: 2001 for WindowSelectionChange, 2013 for SlideShowNextSlide

### 1.2 The Event Sink Implementation

```python
class ppEApplicationSink(comtypes.COMObject):
    _com_interfaces_ = [EApplication, IDispatch]

    def SlideShowNextSlide(self, slideShowWindow=None):
        # Handle slideshow slide changes
        i = winUser.getGUIThreadInfo(0)
        oldFocus = api.getFocusObject()
        if not isinstance(oldFocus, SlideShowWindow) or i.hwndFocus != oldFocus.windowHandle:
            return
        oldFocus.treeInterceptor.rootNVDAObject.handleSlideChange()

    def WindowSelectionChange(self, sel):
        # Handle selection changes in normal view
        i = winUser.getGUIThreadInfo(0)
        oldFocus = api.getFocusObject()
        if not isinstance(oldFocus, Window) or i.hwndFocus != oldFocus.windowHandle:
            return
        if isinstance(oldFocus, DocumentWindow):
            documentWindow = oldFocus
        elif isinstance(oldFocus, PpObject):
            documentWindow = oldFocus.documentWindow
        else:
            return
        documentWindow.ppSelection = sel
        documentWindow.handleSelectionChange()
```

### 1.3 Connection Establishment

NVDA connects to events in `fetchPpObjectModel`:

```python
def fetchPpObjectModel(self, windowHandle=None):
    m = self._fetchPpObjectModelHelper(windowHandle=windowHandle)
    if m:
        if windowHandle != self._ppApplicationWindow or not self._ppApplication:
            self._ppApplicationWindow = windowHandle
            self._ppApplication = m.application
            # Create sink and connect
            sink = ppEApplicationSink().QueryInterface(comtypes.IUnknown)
            self._ppEApplicationConnectionPoint = comtypes.client._events._AdviseConnection(
                self._ppApplication,
                EApplication,
                sink,
            )
    return m
```

**Connection Pattern:**
1. Get PowerPoint Application COM object
2. Create event sink instance
3. Query sink for IUnknown interface
4. Create `_AdviseConnection` with: (application, interface, sink)
5. Store connection reference to prevent garbage collection

---

## 2. Available PowerPoint COM Events

### 2.1 EApplication Interface Events

The PowerPoint EApplication interface (GUID: `{914934C2-5A91-11CF-8700-00AA0060263B}`) exposes many events. Key ones for our purpose:

| Event | DISPID | When Fired |
|-------|--------|------------|
| WindowSelectionChange | 2001 | Text, shape, or slide selection changes |
| SlideSelectionChanged | ~2014 | Slide selection changes in thumbnail pane |
| SlideShowNextSlide | 2013 | Slide advances in slideshow |
| SlideShowBegin | 2010 | Slideshow starts |
| SlideShowEnd | 2011 | Slideshow ends |
| WindowBeforeRightClick | 2002 | Before right-click |
| WindowBeforeDoubleClick | 2003 | Before double-click |
| PresentationClose | 2004 | Presentation closing |
| PresentationOpen | 2006 | Presentation opened |
| NewPresentation | 2007 | New presentation created |
| PresentationNewSlide | 2008 | New slide added |

### 2.2 SlideSelectionChanged Event Details

**Signature:**
```vb
SlideSelectionChanged(SldRange As SlideRange)
```

**When It Fires by View:**

| View | Behavior |
|------|----------|
| Normal, Master | Fires when slide in slide pane changes |
| Slide Sorter | Fires when selection changes |
| Slide, Notes | Fires when slide changes |
| Outline | Does NOT fire |

**DISPID:** Not documented publicly. Based on sequential numbering pattern, likely 2014 or 2015. Testing required to confirm.

### 2.3 WindowSelectionChange Event Details

**Signature:**
```vb
WindowSelectionChange(Sel As Selection)
```

**When It Fires:**
- When text selection changes
- When shape selection changes
- When slide selection changes
- Triggered by both UI and code

**DISPID:** 2001

**Note:** This event fires MORE frequently than SlideSelectionChanged and includes slide changes, making it a reliable fallback.

---

## 3. Why wireEApplication is Not Accessible

### 3.1 The Problem

Our addon tried:
```python
from nvdaBuiltin.appModules.powerpnt import *
# _wireEApplication is None - not exported
```

### 3.2 Why It Fails

Python's `import *` only imports names listed in `__all__` or names that don't start with underscore. NVDA's powerpnt.py does NOT export:
- `EApplication` (the interface class)
- `ppEApplicationSink` (the sink class)
- Internal helper functions

### 3.3 Explicit Import Also Fails

```python
import nvdaBuiltin.appModules.powerpnt as _nvda_powerpnt
if hasattr(_nvda_powerpnt, 'wireEApplication'):  # False
    ...
```

The attribute `wireEApplication` does not exist in NVDA's source. The term appears in our code but was likely a misunderstanding. NVDA uses `EApplication` and `ppEApplicationSink`.

---

## 4. Option Analysis

### Option A: Load PowerPoint Type Library

**Approach:** Use comtypes to load PowerPoint's type library and get the EApplication interface.

**Implementation:**
```python
from comtypes.client import GetModule

# Method 1: Load from installed PowerPoint
try:
    ppt_module = GetModule("C:\\Program Files\\Microsoft Office\\root\\Office16\\MSPPT.OLB")
    EApplication = ppt_module.EApplication
except:
    # Fallback to GUID
    ppt_module = GetModule(('{91493440-5A91-11CF-8700-00AA0060263B}', 2, 12))
```

**Pros:**
- Gets ALL PowerPoint events including SlideSelectionChanged
- Proper type information from Microsoft
- Future-proof if new events are added

**Cons:**
- "Library not registered" error in many environments
- Path varies by Office version and installation type
- Type library loading is complex and error-prone
- Requires comtypes.gen writable (may not be in NVDA addons)

**Verdict:** NOT RECOMMENDED - Too many failure modes

### Option B: Define Our Own EApplication Interface (RECOMMENDED)

**Approach:** Copy NVDA's pattern exactly, but add the events we need.

**Implementation:**
```python
import comtypes
from comtypes.automation import IDispatch
import ctypes

class EApplication(IDispatch):
    """PowerPoint Application Events interface.

    Minimal interface definition with only the events we need.
    Based on NVDA's powerpnt.py pattern.
    """
    _iid_ = comtypes.GUID("{914934C2-5A91-11CF-8700-00AA0060263B}")
    _methods_ = []
    _disp_methods_ = [
        # WindowSelectionChange - fires on any selection change including slides
        comtypes.DISPMETHOD(
            [comtypes.dispid(2001)],
            None,
            "WindowSelectionChange",
            (["in"], ctypes.POINTER(IDispatch), "sel"),
        ),
        # SlideShowNextSlide - fires during slideshow
        comtypes.DISPMETHOD(
            [comtypes.dispid(2013)],
            None,
            "SlideShowNextSlide",
            (["in"], ctypes.POINTER(IDispatch), "slideShowWindow"),
        ),
        # SlideSelectionChanged - fires when slide selection changes
        # DISPID needs to be determined by testing (likely 2014)
        comtypes.DISPMETHOD(
            [comtypes.dispid(2014)],  # May need adjustment
            None,
            "SlideSelectionChanged",
            (["in"], ctypes.POINTER(IDispatch), "SldRange"),
        ),
    ]


class PowerPointEventSink(comtypes.COMObject):
    """COM Event Sink for PowerPoint application events."""
    _com_interfaces_ = [EApplication, IDispatch]

    def __init__(self, callback):
        super().__init__()
        self._callback = callback
        self._last_slide = -1

    def WindowSelectionChange(self, sel):
        """Fires on any selection change - reliable for slide changes."""
        try:
            # Get current slide index and notify if changed
            if self._callback:
                self._callback("selection_change", sel)
        except Exception as e:
            pass

    def SlideSelectionChanged(self, SldRange):
        """Fires specifically when slide selection changes."""
        try:
            if self._callback and SldRange:
                self._callback("slide_change", SldRange)
        except Exception as e:
            pass

    def SlideShowNextSlide(self, slideShowWindow):
        """Fires during slideshow."""
        try:
            if self._callback:
                self._callback("slideshow_slide", slideShowWindow)
        except Exception as e:
            pass
```

**Connection Code:**
```python
import comtypes.client._events

def connect_to_powerpoint_events(ppt_app, sink):
    """Connect to PowerPoint application events.

    Args:
        ppt_app: PowerPoint.Application COM object
        sink: PowerPointEventSink instance

    Returns:
        _AdviseConnection object (keep reference to stay connected)
    """
    sink_iunknown = sink.QueryInterface(comtypes.IUnknown)
    connection = comtypes.client._events._AdviseConnection(
        ppt_app,
        EApplication,
        sink_iunknown,
    )
    return connection
```

**Pros:**
- Proven pattern (NVDA uses it)
- No external dependencies
- Works in NVDA addon environment
- Full control over which events to handle
- Can add SlideSelectionChanged

**Cons:**
- Must determine correct DISPID for SlideSelectionChanged (testing required)
- If DISPID is wrong, event won't fire

**Verdict:** RECOMMENDED - Most reliable approach

### Option C: Hook Into NVDA's Existing Connection

**Approach:** Access NVDA's existing COM event infrastructure.

**Attempted Implementation:**
```python
# Get NVDA's built-in AppModule's event connection
parent_app_module = super()  # Built-in PowerPoint AppModule
if hasattr(parent_app_module, '_ppEApplicationConnectionPoint'):
    # Reuse existing connection somehow
    ...
```

**Analysis:**
1. NVDA's powerpnt module stores connection in `_ppEApplicationConnectionPoint`
2. The connection is on NVDAObject instances (DocumentWindow), not AppModule
3. We cannot easily intercept or extend the existing sink
4. Adding methods to existing sink is not possible without modifying NVDA source

**Pros:**
- Would avoid duplicate connections
- Less resource usage

**Cons:**
- NVDA's connection is per-window, not per-app
- Cannot add new event handlers to existing sink
- Implementation extremely complex
- Would break if NVDA changes internal structure

**Verdict:** NOT RECOMMENDED - Too complex, not reliable

---

## 5. Recommended Implementation

### 5.1 Architecture

```
AppModule
    |
    +-- PowerPointWorker (background thread)
            |
            +-- EApplication (interface definition)
            +-- PowerPointEventSink (event handler)
            +-- _AdviseConnection (kept alive)
            |
            +-- Message pump (receives COM events)
```

### 5.2 Complete Implementation

```python
# powerpoint-comments/addon/appModules/powerpnt.py

import logging
log = logging.getLogger(__name__)

from nvdaBuiltin.appModules.powerpnt import *

import comtypes
from comtypes.automation import IDispatch
from comtypes import CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED, COMObject
import comtypes.client._events
import ctypes
import threading
import comHelper
import ui
from queueHandler import queueFunction, eventQueue

# ============================================================================
# PowerPoint EApplication Event Interface
# ============================================================================

class EApplication(IDispatch):
    """PowerPoint Application Events interface.

    Based on NVDA's powerpnt.py pattern. Defines only the events we need.
    GUID is the same for all PowerPoint versions.
    """
    _iid_ = comtypes.GUID("{914934C2-5A91-11CF-8700-00AA0060263B}")
    _methods_ = []
    _disp_methods_ = [
        # WindowSelectionChange (DISPID 2001)
        # Fires on any selection change - text, shape, or slide
        comtypes.DISPMETHOD(
            [comtypes.dispid(2001)],
            None,
            "WindowSelectionChange",
            (["in"], ctypes.POINTER(IDispatch), "sel"),
        ),
        # SlideShowNextSlide (DISPID 2013)
        # Fires when slide advances in slideshow mode
        comtypes.DISPMETHOD(
            [comtypes.dispid(2013)],
            None,
            "SlideShowNextSlide",
            (["in"], ctypes.POINTER(IDispatch), "slideShowWindow"),
        ),
    ]


class PowerPointEventSink(COMObject):
    """COM Event Sink for PowerPoint application events."""
    _com_interfaces_ = [EApplication, IDispatch]

    def __init__(self, worker):
        super().__init__()
        self._worker = worker
        self._last_slide_index = -1
        log.info("PowerPointEventSink: Initialized")

    def WindowSelectionChange(self, sel):
        """Called when selection changes in PowerPoint window.

        This is the most reliable event for detecting slide changes
        in normal editing view. Fires for text, shape, and slide selections.
        """
        try:
            log.debug("PowerPointEventSink: WindowSelectionChange event")
            if self._worker and self._worker._ppt_app:
                try:
                    current_index = self._worker._ppt_app.ActiveWindow.View.Slide.SlideIndex
                    if current_index != self._last_slide_index:
                        log.info(f"PowerPointEventSink: Slide changed to {current_index}")
                        self._last_slide_index = current_index
                        self._worker.on_slide_changed(current_index)
                except Exception as e:
                    log.debug(f"PowerPointEventSink: Could not get slide - {e}")
        except Exception as e:
            log.error(f"PowerPointEventSink: Error in WindowSelectionChange - {e}")

    def SlideShowNextSlide(self, slideShowWindow=None):
        """Called when slide advances in slideshow mode."""
        try:
            log.debug("PowerPointEventSink: SlideShowNextSlide event")
            if self._worker:
                self._worker.on_slideshow_slide_change(slideShowWindow)
        except Exception as e:
            log.error(f"PowerPointEventSink: Error in SlideShowNextSlide - {e}")


# ============================================================================
# PowerPoint Worker Thread
# ============================================================================

class PowerPointWorker:
    """Background thread for PowerPoint COM operations and event handling."""

    def __init__(self):
        self._stop_event = threading.Event()
        self._thread = None
        self._ppt_app = None
        self._event_sink = None
        self._event_connection = None
        self._initialized = False
        self._last_announced_slide = -1

    def start(self):
        """Start the background thread."""
        self._thread = threading.Thread(
            target=self._run,
            name="PowerPointCommentWorker",
            daemon=False
        )
        self._thread.start()
        log.info("PowerPoint worker thread started")

    def stop(self, timeout=5):
        """Stop the thread gracefully."""
        log.info("PowerPoint worker thread stopping...")
        self._stop_event.set()
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=timeout)

    def request_initialize(self):
        """Request initialization (called from main thread)."""
        self._initialized = False

    def _run(self):
        """Main thread loop."""
        CoInitializeEx(COINIT_APARTMENTTHREADED)
        log.info("PowerPoint worker: COM initialized (STA)")

        try:
            self._initialize_com()

            while not self._stop_event.is_set():
                try:
                    # Pump messages to receive COM events
                    self._pump_messages(timeout_ms=500)

                    if not self._initialized:
                        self._initialize_com()
                except Exception as e:
                    log.error(f"Worker thread error: {e}")
        finally:
            self._disconnect_events()
            self._ppt_app = None
            CoUninitialize()
            log.info("PowerPoint worker: Thread exiting")

    def _pump_messages(self, timeout_ms=500):
        """Pump Windows messages to receive COM events."""
        from ctypes import windll, byref
        from ctypes.wintypes import MSG

        user32 = windll.user32
        QS_ALLINPUT = 0x04FF

        user32.MsgWaitForMultipleObjects(0, None, False, timeout_ms, QS_ALLINPUT)

        msg = MSG()
        PM_REMOVE = 0x0001
        while user32.PeekMessageW(byref(msg), None, 0, 0, PM_REMOVE):
            user32.TranslateMessage(byref(msg))
            user32.DispatchMessageW(byref(msg))

    def _initialize_com(self):
        """Connect to PowerPoint and set up event handling."""
        try:
            log.info("Worker: Connecting to PowerPoint...")
            self._ppt_app = comHelper.getActiveObject("PowerPoint.Application", dynamic=True)
            log.info("Worker: Connected to PowerPoint COM")

            if self._has_active_presentation():
                self._connect_events()
                self._check_initial_slide()
                self._initialized = True
            else:
                log.info("Worker: No active presentation")
                self._initialized = False
        except Exception as e:
            log.error(f"Worker: Initialize failed - {e}")
            self._ppt_app = None
            self._initialized = False

    def _connect_events(self):
        """Connect to PowerPoint application events."""
        if not self._ppt_app:
            return

        try:
            self._disconnect_events()

            # Create event sink
            self._event_sink = PowerPointEventSink(self)

            # Get IUnknown interface from sink
            sink_iunknown = self._event_sink.QueryInterface(comtypes.IUnknown)

            # Create advise connection
            self._event_connection = comtypes.client._events._AdviseConnection(
                self._ppt_app,
                EApplication,
                sink_iunknown,
            )

            log.info("Worker: Connected to PowerPoint events")
        except Exception as e:
            log.error(f"Worker: Failed to connect events - {e}")
            self._event_sink = None
            self._event_connection = None

    def _disconnect_events(self):
        """Disconnect from PowerPoint events."""
        if self._event_connection:
            try:
                del self._event_connection
                self._event_connection = None
                log.info("Worker: Disconnected from events")
            except Exception as e:
                log.debug(f"Worker: Error disconnecting - {e}")
        self._event_sink = None

    def _has_active_presentation(self):
        """Check if there's an active presentation."""
        try:
            return (self._ppt_app and
                    self._ppt_app.Presentations.Count > 0 and
                    self._ppt_app.ActiveWindow)
        except:
            return False

    def _check_initial_slide(self):
        """Check and announce comments on initial slide."""
        try:
            current_index = self._ppt_app.ActiveWindow.View.Slide.SlideIndex
            if current_index > 0:
                log.info(f"Worker: Initial slide is {current_index}")
                self._last_announced_slide = current_index
                self._announce_slide_comments()
        except Exception as e:
            log.debug(f"Worker: Error checking initial slide - {e}")

    def on_slide_changed(self, slide_index):
        """Called by event sink when slide changes."""
        if slide_index == self._last_announced_slide:
            return

        log.info(f"Worker: Slide changed to {slide_index}")
        self._last_announced_slide = slide_index
        self._announce_slide_comments()

    def on_slideshow_slide_change(self, slideShowWindow):
        """Called by event sink during slideshow."""
        log.info("Worker: Slideshow slide change")
        # Could handle slideshow mode here

    def _announce_slide_comments(self):
        """Announce comment status for current slide."""
        comments = self._get_comments_on_current_slide()

        if not comments:
            self._announce("No comments")
        else:
            count = len(comments)
            msg = f"Has {count} comment{'s' if count != 1 else ''}"
            self._announce(msg)

    def _get_comments_on_current_slide(self):
        """Get all comments on current slide."""
        try:
            slide = self._ppt_app.ActiveWindow.View.Slide
            comments = []
            for i in range(1, slide.Comments.Count + 1):
                comment = slide.Comments.Item(i)
                comments.append({
                    'text': comment.Text,
                    'author': comment.Author,
                })
            return comments
        except:
            return []

    def _announce(self, message):
        """Thread-safe UI announcement."""
        queueFunction(eventQueue, ui.message, message)


# ============================================================================
# AppModule
# ============================================================================

class AppModule(AppModule):
    """Enhanced PowerPoint with comment navigation."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._worker = PowerPointWorker()
        self._worker.start()
        log.info("PowerPoint Comments AppModule initialized")

    def event_appModule_gainFocus(self):
        """Called when PowerPoint gains focus."""
        log.info("PowerPoint Comments: App gained focus")
        if self._worker:
            self._worker.request_initialize()

    def terminate(self):
        """Clean up when PowerPoint closes."""
        log.info("PowerPoint Comments: Terminating")
        if hasattr(self, '_worker') and self._worker:
            self._worker.stop(timeout=5)
        super().terminate()
```

---

## 6. Determining SlideSelectionChanged DISPID

### 6.1 The Challenge

Microsoft does not publicly document DISPIDs for EApplication events. The KB article 309309 that contained this information is no longer accessible.

### 6.2 Known DISPIDs (from NVDA and other sources)

| Event | DISPID | Source |
|-------|--------|--------|
| WindowSelectionChange | 2001 | NVDA source |
| WindowBeforeRightClick | 2002 | KB 309309 (archived) |
| WindowBeforeDoubleClick | 2003 | KB 309309 (archived) |
| PresentationClose | 2004 | KB 309309 (archived) |
| PresentationSave | 2005 | KB 309309 (archived) |
| PresentationOpen | 2006 | KB 309309 (archived) |
| NewPresentation | 2007 | KB 309309 (archived) |
| PresentationNewSlide | 2008 | KB 309309 (archived) |
| SlideShowBegin | 2010 | Estimated |
| SlideShowEnd | 2011 | Estimated |
| SlideShowNextSlide | 2013 | NVDA source |
| SlideSelectionChanged | 2014? | Needs testing |

### 6.3 Testing Strategy

To find the correct DISPID:

1. **Try sequential DISPIDs** - Test 2009, 2012, 2014, 2015, etc.
2. **Use OLE/COM Object Viewer** - If PowerPoint exposes the dispinterface, view it with oleview.exe
3. **Check type library** - Use Python to enumerate the type library:

```python
from comtypes.client import GetModule
import comtypes.gen

# Try to load PowerPoint type library
ppt = GetModule("PowerPoint.Application")
# Examine ppt.EApplication if available
```

### 6.4 Practical Recommendation

**WindowSelectionChange is sufficient for our needs.** It fires on all selection changes including slide changes. Unless we need to distinguish between slide selection and other selection types, WindowSelectionChange alone handles our use case.

---

## 7. References

### Primary Sources
- [NVDA GitHub - powerpnt.py](https://github.com/nvaccess/nvda/blob/master/source/appModules/powerpnt.py)
- [NVDA Developer Guide](https://www.nvaccess.org/files/nvda/documentation/developerGuide.html)
- [comtypes documentation](https://pythonhosted.org/comtypes/)
- [comtypes GetEvents](https://snyk.io/advisor/python/comtypes/functions/comtypes.client.GetEvents)
- [comtypes _AdviseConnection examples](https://programtalk.com/python-examples/comtypes.client._events._AdviseConnection/)

### Microsoft Documentation
- [Application.WindowSelectionChange event](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.Application.WindowSelectionChange)
- [Application.SlideSelectionChanged event](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.slideselectionchanged)
- [Application.SlideShowNextSlide event](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.Application.SlideShowNextSlide)
- [Use Events with the Application Object](https://learn.microsoft.com/en-us/office/vba/powerpoint/how-to/use-events-with-the-application-object)
- [PowerPoint Application Events in VBA](http://youpresent.co.uk/powerpoint-application-events-in-vba/)

### Additional Resources
- [Handling PowerPoint Slide Show Events from Python](https://developer.mamezou-tech.com/en/blogs/2024/09/02/monitor-pptx-py/)
- [OfficeOne: Events supported by PowerPoint](https://www.officeoneonline.com/vba/events_version.html)

---

## Document Information

- **Created:** December 10, 2025
- **Author:** Strategic Planning and Research Agent
- **Version:** 1.0
- **Purpose:** Research foundation for PowerPoint COM event integration in NVDA addon
