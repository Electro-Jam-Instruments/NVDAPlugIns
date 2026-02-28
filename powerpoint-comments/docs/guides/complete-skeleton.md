# Complete Skeleton - NVDA PowerPoint Addon

A complete working example showing all the pieces together. Use this as a reference or starting point.

## File Structure

### Linear Walkthrough

**powerpoint-comments/ folder contains:**
- **addon/** - The addon package
  - **manifest.ini** - Addon metadata
  - **appModules/powerpnt.py** - PowerPoint module
- **buildVars.py** - Build configuration
- **sconstruct** - Scons build script

### 2D Visual Map

```
powerpoint-comments/
├── addon/
│   ├── manifest.ini
│   └── appModules/
│       └── powerpnt.py
├── buildVars.py
└── sconstruct
```

## manifest.ini (Complete Example)

```ini
name = powerPointComments
summary = "PowerPoint Comments NVDA Addon"
description = """Make PowerPoint comments accessible with NVDA screen reader.
Adds comment count announcements, notes detection, and keyboard shortcuts."""
author = "Your Name"
url = https://github.com/your-org/your-repo
version = 0.1.0
docFileName = readme.html
minimumNVDAVersion = 2024.1
lastTestedNVDAVersion = 2025.3
```

**Quoting rules:**
- `name` - no quotes (single word)
- `summary` - double quotes (has spaces)
- `description` - triple quotes (multiline)
- `version`, `url` - no quotes

## buildVars.py (Complete Example)

```python
# -*- coding: UTF-8 -*-

# Build customizations
# Change this file instead of sconstruct or manifest files

# Full geance with samples can be found at:
# https://github.com/nvaccess/addon-template/blob/master/buildVars.py

addon_info = {
    "addon_name": "powerPointComments",
    "addon_summary": "PowerPoint Comments NVDA Addon",
    "addon_description": """Make PowerPoint comments accessible with NVDA.""",
    "addon_version": "0.1.0",
    "addon_author": "Your Name <your@email.com>",
    "addon_url": "https://github.com/your-org/your-repo",
    "addon_docFileName": "readme.html",
    "addon_minimumNVDAVersion": "2024.1",
    "addon_lastTestedNVDAVersion": "2025.3",
}

# Define the python files that are the sources of your add-on.
pythonSources = []

# Files that contain strings for translation
i18nSources = []

# Files that will be ignored when building the nvda-addon file
excludedFiles = []

# Base language for the NVDA add-on
baseLanguage = "en"
```

## powerpnt.py (Complete Skeleton)

```python
# appModules/powerpnt.py
# Complete skeleton for NVDA PowerPoint addon

# Version - update in manifest.ini and buildVars.py too
ADDON_VERSION = "0.1.0"

# ============================================================================
# IMPORTS - Order matters!
# ============================================================================

# Logging FIRST
import logging
log = logging.getLogger(__name__)
log.info(f"PowerPoint addon loading (v{ADDON_VERSION})")

# CRITICAL: Import EVERYTHING from built-in PowerPoint module
# This is the ONLY pattern that works for extending built-in appModules
from nvdaBuiltin.appModules.powerpnt import *
log.info("Built-in powerpnt imported successfully")

# Additional imports
import comHelper  # NVDA's COM helper - handles UIAccess privileges
import ui
import api
import speech
import threading
import ctypes
import comtypes
from comtypes import CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED, COMObject, GUID
from comtypes.automation import IDispatch
from comtypes.client._events import _AdviseConnection
from queueHandler import queueFunction, eventQueue
from scriptHandler import script


# ============================================================================
# COM EVENT INTERFACE - Defined locally (NOT from type library)
# ============================================================================

class EApplication(IDispatch):
    """PowerPoint Application Events interface.

    CRITICAL: Define locally - type library loading fails with "Library not registered"
    GUID: {914934C2-5A91-11CF-8700-00AA0060263B}
    """
    _iid_ = GUID("{914934C2-5A91-11CF-8700-00AA0060263B}")
    _methods_ = []
    _disp_methods_ = [
        # WindowSelectionChange (DISPID 2001) - edit mode slide changes
        comtypes.DISPMETHOD(
            [comtypes.dispid(2001)], None, "WindowSelectionChange",
            (["in"], ctypes.POINTER(IDispatch), "sel"),
        ),
        # SlideShowBegin (DISPID 2010)
        comtypes.DISPMETHOD(
            [comtypes.dispid(2010)], None, "SlideShowBegin",
            (["in"], ctypes.POINTER(IDispatch), "wn"),
        ),
        # SlideShowEnd (DISPID 2012)
        comtypes.DISPMETHOD(
            [comtypes.dispid(2012)], None, "SlideShowEnd",
            (["in"], ctypes.POINTER(IDispatch), "pres"),
        ),
        # SlideShowNextSlide (DISPID 2013) - slideshow slide changes
        comtypes.DISPMETHOD(
            [comtypes.dispid(2013)], None, "SlideShowNextSlide",
            (["in"], ctypes.POINTER(IDispatch), "slideShowWindow"),
        ),
    ]


# ============================================================================
# EVENT SINK - Receives COM events
# ============================================================================

class PowerPointEventSink(COMObject):
    """COM Event Sink for PowerPoint application events."""

    _com_interfaces_ = [EApplication, IDispatch]

    def __init__(self, worker):
        super().__init__()
        self._worker = worker
        self._last_slide_index = -1
        log.info("PowerPointEventSink initialized")

    def WindowSelectionChange(self, sel):
        """Fires on slide/shape/text selection change in edit mode."""
        try:
            if self._worker and self._worker._ppt_app:
                # Get window from selection's parent (handles multiple presentations)
                window = None
                try:
                    window = sel.Parent
                except:
                    window = self._worker._ppt_app.ActiveWindow

                if window:
                    current_index = window.View.Slide.SlideIndex
                    if current_index != self._last_slide_index:
                        log.info(f"Slide changed to {current_index}")
                        self._last_slide_index = current_index
                        self._worker.on_slide_changed(current_index, window)
        except Exception as e:
            log.debug(f"WindowSelectionChange error: {e}")

    def SlideShowBegin(self, wn):
        """Fires when slideshow starts."""
        log.info("SlideShowBegin event")
        if self._worker:
            self._worker.on_slideshow_begin(wn)

    def SlideShowEnd(self, pres):
        """Fires when slideshow ends."""
        log.info("SlideShowEnd event")
        if self._worker:
            self._worker.on_slideshow_end(pres)

    def SlideShowNextSlide(self, slideShowWindow):
        """Fires when slide advances in slideshow (NOT for first slide!)."""
        try:
            if self._worker and slideShowWindow:
                slide_index = slideShowWindow.View.Slide.SlideIndex
                if slide_index != self._last_slide_index:
                    log.info(f"Slideshow slide changed to {slide_index}")
                    self._last_slide_index = slide_index
                    self._worker.on_slideshow_slide_changed(slide_index, slideShowWindow)
        except Exception as e:
            log.debug(f"SlideShowNextSlide error: {e}")


# ============================================================================
# WORKER THREAD - Handles all COM operations
# ============================================================================

class PowerPointWorker:
    """Background thread for PowerPoint COM operations.

    All COM work runs here with proper STA initialization.
    Main thread communicates via request_* methods.
    Worker communicates back via queueFunction(eventQueue, ...).
    """

    def __init__(self):
        self._stop_event = threading.Event()
        self._thread = None
        self._ppt_app = None
        self._event_sink = None
        self._event_connection = None  # MUST keep reference alive!
        self._initialized = False
        self._has_focus = False

        # Cached data for main thread to read
        self._last_slide_index = -1
        self._comment_count = 0
        self._has_notes = False
        self._in_slideshow = False

    # ---- Lifecycle ----

    def start(self):
        """Start the worker thread."""
        self._thread = threading.Thread(
            target=self._run,
            name="PowerPointWorker",
            daemon=False  # Non-daemon for clean shutdown
        )
        self._thread.start()
        log.info("Worker thread started")

    def stop(self, timeout=5):
        """Stop the worker thread."""
        log.info("Worker thread stopping...")
        self._stop_event.set()
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=timeout)
        log.info("Worker thread stopped")

    # ---- Main thread requests (non-blocking) ----

    def request_initialize(self):
        """Called from event_appModule_gainFocus - signals worker to init."""
        log.info("Initialize requested")
        self._has_focus = True
        self._initialized = False  # Force reinit

    def request_read_notes(self):
        """Called from keyboard script - signals worker to read notes."""
        self._read_notes_pending = True

    # ---- Worker thread main loop ----

    def _run(self):
        """Main worker thread loop."""
        # Initialize COM in STA mode (required for Office)
        CoInitializeEx(COINIT_APARTMENTTHREADED)
        log.info("COM initialized (STA)")

        try:
            while not self._stop_event.is_set():
                # Pump Windows messages (REQUIRED for COM events)
                self._pump_messages(timeout_ms=500)

                # Initialize if needed
                if not self._initialized and self._has_focus:
                    self._initialize_com()

        finally:
            self._disconnect_events()
            self._ppt_app = None
            CoUninitialize()
            log.info("COM uninitialized")

    def _pump_messages(self, timeout_ms=500):
        """Pump Windows messages to receive COM events."""
        from ctypes import windll, byref
        from ctypes.wintypes import MSG

        user32 = windll.user32
        user32.MsgWaitForMultipleObjects(0, None, False, timeout_ms, 0x04FF)

        msg = MSG()
        while user32.PeekMessageW(byref(msg), None, 0, 0, 0x0001):
            user32.TranslateMessage(byref(msg))
            user32.DispatchMessageW(byref(msg))

    # ---- COM initialization ----

    def _initialize_com(self):
        """Connect to PowerPoint and set up events."""
        try:
            # CRITICAL: Use comHelper, NOT GetActiveObject
            self._ppt_app = comHelper.getActiveObject(
                "PowerPoint.Application",
                dynamic=True
            )
            log.info("Connected to PowerPoint")

            # Check for active presentation
            if self._ppt_app.Presentations.Count > 0:
                self._connect_events()
                self._initialized = True
                # Announce initial state
                self._announce_current_slide()

        except Exception as e:
            log.info(f"PowerPoint not available: {e}")
            self._ppt_app = None

    def _connect_events(self):
        """Connect to PowerPoint COM events."""
        if not self._ppt_app:
            return

        try:
            self._disconnect_events()

            # Create sink
            self._event_sink = PowerPointEventSink(self)

            # Get IUnknown
            sink_iunknown = self._event_sink.QueryInterface(comtypes.IUnknown)

            # Connect - KEEP THIS REFERENCE!
            self._event_connection = _AdviseConnection(
                self._ppt_app,
                EApplication,
                sink_iunknown
            )
            log.info("Connected to PowerPoint events")

        except Exception as e:
            log.error(f"Failed to connect events: {e}")

    def _disconnect_events(self):
        """Disconnect from PowerPoint events."""
        if self._event_connection:
            del self._event_connection
            self._event_connection = None
        self._event_sink = None

    # ---- Event handlers (called by sink) ----

    def on_slide_changed(self, slide_index, window):
        """Called when slide changes in edit mode."""
        self._last_slide_index = slide_index
        self._announce_current_slide()

    def on_slideshow_begin(self, wn):
        """Called when slideshow starts."""
        self._in_slideshow = True
        # Cache first slide (SlideShowNextSlide doesn't fire for slide 1)
        self._cache_slideshow_data(wn)

    def on_slideshow_end(self, pres):
        """Called when slideshow ends."""
        self._in_slideshow = False

    def on_slideshow_slide_changed(self, slide_index, slideshow_window):
        """Called when slide changes in slideshow."""
        self._last_slide_index = slide_index
        self._cache_slideshow_data(slideshow_window)

    # ---- COM queries ----

    def _get_current_slide(self):
        """Get current slide COM object."""
        try:
            return self._ppt_app.ActiveWindow.View.Slide
        except:
            return None

    def _get_comment_count(self, slide):
        """Get comment count for slide."""
        try:
            return slide.Comments.Count
        except:
            return 0

    def _has_meeting_notes(self, slide):
        """Check if slide has meeting notes (contains ****)."""
        try:
            notes_page = slide.NotesPage
            placeholder = notes_page.Shapes.Placeholders(2)
            if placeholder.HasTextFrame:
                text = placeholder.TextFrame.TextRange.Text
                return '****' in text
        except:
            pass
        return False

    def _cache_slideshow_data(self, slideshow_window):
        """Cache data for slideshow mode."""
        try:
            slide = slideshow_window.View.Slide
            self._comment_count = self._get_comment_count(slide)
            self._has_notes = self._has_meeting_notes(slide)
        except Exception as e:
            log.debug(f"Error caching slideshow data: {e}")

    # ---- Announcements ----

    def _announce_current_slide(self):
        """Announce current slide status."""
        slide = self._get_current_slide()
        if not slide:
            return

        count = self._get_comment_count(slide)
        has_notes = self._has_meeting_notes(slide)

        # Cache for main thread
        self._comment_count = count
        self._has_notes = has_notes

        # Build announcement
        parts = []
        if has_notes:
            parts.append("has notes")
        if count > 0:
            parts.append(f"Has {count} comment{'s' if count != 1 else ''}")

        if parts:
            self._announce(", ".join(parts))

    def _announce(self, message):
        """Thread-safe announcement."""
        queueFunction(eventQueue, ui.message, message)


# ============================================================================
# APP MODULE
# ============================================================================

class AppModule(AppModule):  # Inherits from just-imported AppModule!
    """Extended PowerPoint support with comment navigation."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        log.info("PowerPoint Comments AppModule initializing")

        # Start worker thread
        self._worker = PowerPointWorker()
        self._worker.start()

    def terminate(self):
        """Clean shutdown."""
        log.info("PowerPoint Comments AppModule terminating")
        if self._worker:
            self._worker.stop(timeout=5)
        super().terminate()

    def event_appModule_gainFocus(self):
        """Called when PowerPoint gains focus.

        CRITICAL:
        - Do NOT call super() - this method doesn't exist in parent
        - Do NOT do heavy work - delegate to worker thread
        """
        log.info("PowerPoint gained focus")
        if self._worker:
            self._worker.request_initialize()

    # ---- Keyboard Scripts ----

    @script(
        description="Read slide notes",
        gesture="kb:control+alt+n",
        category="PowerPoint Comments"
    )
    def script_readNotes(self, gesture):
        """Read the current slide's notes."""
        if self._worker:
            self._worker.request_read_notes()


# Log successful load
log.info(f"PowerPoint Comments addon loaded successfully (v{ADDON_VERSION})")
```

## How to Build

1. Install scons: `pip install scons`
2. Clone NVDA addon template: https://github.com/nvaccess/addon-template
3. Copy your addon files into the template structure
4. Run: `scons`
5. Output: `powerPointComments-0.1.0.nvda-addon`

## How to Test

1. Copy `powerpnt.py` to `%APPDATA%\nvda\scratchpad\appModules\`
2. Enable scratchpad: NVDA Settings > Advanced > Developer Scratchpad
3. Reload: NVDA+Ctrl+F3
4. Check log: NVDA+F1

## Key Patterns Summary

| Pattern | Implementation |
|---------|----------------|
| Inheritance | `from nvdaBuiltin...import *` then `class AppModule(AppModule)` |
| COM access | `comHelper.getActiveObject(..., dynamic=True)` |
| COM events | Define EApplication locally, use `_AdviseConnection` |
| Threading | Worker thread with `CoInitializeEx(COINIT_APARTMENTTHREADED)` |
| Event handlers | No super() on `event_appModule_gainFocus`, delegate to worker |
| UI updates | `queueFunction(eventQueue, ui.message, text)` |
| Event connection | MUST keep `_event_connection` reference alive |

## Window Class Names Reference

| Window Class | Object Type |
|--------------|-------------|
| `mdiClass` | Main slide editing area |
| `paneClassDC` | Slide thumbnails pane |
| `NetUIHWND` | Ribbon, task panes |
| `screenClass` | Slideshow window |

## Detecting Modes

```python
# Check if in slideshow mode
def is_in_slideshow(self):
    try:
        return self._ppt_app.SlideShowWindows.Count > 0
    except:
        return False

# Check view type (edit mode)
PP_VIEW_NORMAL = 9
PP_VIEW_SLIDESHOW = 1

def get_view_type(self):
    try:
        return self._ppt_app.ActiveWindow.ViewType
    except:
        return None
```
