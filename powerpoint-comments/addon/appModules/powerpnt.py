# appModules/powerpnt.py
# PowerPoint Comments NVDA Addon - Phase 2: Slide Change Detection via COM Events
#
# This module extends NVDA's built-in PowerPoint support to add
# comment navigation features.
#
# Pattern: NVDA Developer Guide - extending built-in appModules
# https://download.nvaccess.org/documentation/developerGuide.html
# Uses: from nvdaBuiltin.appModules.xxx import * then class AppModule(AppModule)

# Addon version - update this and manifest.ini together
ADDON_VERSION = "0.0.63"

# Import logging FIRST so we can log any import issues
import logging
log = logging.getLogger(__name__)
log.info(f"PowerPoint Comments addon: Module loading (v{ADDON_VERSION})")

# Import EVERYTHING from built-in PowerPoint module
# This is the NVDA-documented pattern for extending built-in appModules
from nvdaBuiltin.appModules.powerpnt import *
log.info("PowerPoint Comments addon: Built-in powerpnt imported successfully")

# Additional imports for our functionality
import comHelper  # NVDA's COM helper - handles UIAccess privilege issues
import speech  # For canceling queued speech (v0.0.37)
import ui
import api
import threading
import ctypes
import comtypes
from comtypes import CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED, COMObject, GUID
from comtypes.automation import IDispatch
from comtypes.client._events import _AdviseConnection
from queueHandler import queueFunction, eventQueue
from scriptHandler import script
import inputCore

# ============================================================================
# COM Event Interface - Defined Locally (v0.0.21)
# ============================================================================
#
# CRITICAL: We define the EApplication interface locally because:
# 1. PowerPoint's type library fails to load ("Library not registered")
# 2. This is exactly how NVDA's built-in powerpnt.py does it
# 3. It's reliable and doesn't depend on system type library registration
#
# v0.0.22: Use sel.Parent to get correct window when multiple presentations open
#
# See: .agent/experts/nvda-plugins/research/PowerPoint-COM-Events-Research.md


class EApplication(IDispatch):
    """PowerPoint Application Events interface.

    Defined locally to avoid type library loading issues.
    GUID and DISPIDs match PowerPoint's EApplication interface.

    Interface GUID: {914934C2-5A91-11CF-8700-00AA0060263B}
    """
    _iid_ = GUID("{914934C2-5A91-11CF-8700-00AA0060263B}")
    _methods_ = []
    _disp_methods_ = [
        # WindowSelectionChange (DISPID 2001) - fires on ANY selection change
        # This is the most reliable event for detecting slide changes in edit mode
        comtypes.DISPMETHOD(
            [comtypes.dispid(2001)],
            None,
            "WindowSelectionChange",
            (["in"], ctypes.POINTER(IDispatch), "sel"),
        ),
        # SlideShowBegin (DISPID 2010) - fires when slideshow starts
        comtypes.DISPMETHOD(
            [comtypes.dispid(2010)],
            None,
            "SlideShowBegin",
            (["in"], ctypes.POINTER(IDispatch), "wn"),
        ),
        # SlideShowEnd (DISPID 2012) - fires when slideshow ends
        comtypes.DISPMETHOD(
            [comtypes.dispid(2012)],
            None,
            "SlideShowEnd",
            (["in"], ctypes.POINTER(IDispatch), "pres"),
        ),
        # SlideShowNextSlide (DISPID 2013) - fires during slideshow
        comtypes.DISPMETHOD(
            [comtypes.dispid(2013)],
            None,
            "SlideShowNextSlide",
            (["in"], ctypes.POINTER(IDispatch), "slideShowWindow"),
        ),
    ]


class PowerPointEventSink(COMObject):
    """COM Event Sink for PowerPoint application events.

    Receives WindowSelectionChange events from PowerPoint.
    Calls back to the worker thread to process slide changes.

    v0.0.21: Uses locally-defined EApplication interface instead of type library.
    """

    _com_interfaces_ = [EApplication, IDispatch]

    def __init__(self, worker):
        """Initialize the event sink.

        Args:
            worker: PowerPointWorker instance to notify on events
        """
        super().__init__()
        self._worker = worker
        self._last_slide_index = -1
        log.info("PowerPointEventSink: Initialized with local EApplication interface")

    def WindowSelectionChange(self, sel):
        """Called when selection changes in PowerPoint window.

        This event fires for text, shape, and slide selections.
        We check if the slide index has changed.

        v0.0.22: Use sel.Parent to get the SPECIFIC window that triggered the event.
        This fixes wrong data when multiple presentations are open.

        Args:
            sel: Selection object (IDispatch) - sel.Parent returns the DocumentWindow
        """
        try:
            log.debug("PowerPointEventSink: WindowSelectionChange event received")
            if self._worker and self._worker._ppt_app:
                try:
                    # v0.0.22: Get the SPECIFIC window from sel.Parent
                    # This is the key fix for multiple presentations
                    window = None
                    try:
                        window = sel.Parent
                        log.debug("PowerPointEventSink: Got window from sel.Parent")
                    except Exception as e:
                        log.debug(f"PowerPointEventSink: sel.Parent failed ({e}), using ActiveWindow")
                        window = self._worker._ppt_app.ActiveWindow

                    if window:
                        current_index = window.View.Slide.SlideIndex
                        if current_index != self._last_slide_index:
                            log.info(f"PowerPointEventSink: Slide changed to {current_index}")
                            self._last_slide_index = current_index
                            # Pass the specific window to the worker
                            self._worker.on_slide_changed_event(current_index, window)
                except Exception as e:
                    log.debug(f"PowerPointEventSink: Could not get slide - {e}")
        except Exception as e:
            log.error(f"PowerPointEventSink: Error in WindowSelectionChange - {e}")

    def SlideShowBegin(self, wn):
        """Called when slideshow starts.

        v0.0.56: Track slideshow state to modify announcements.

        Args:
            wn: SlideShowWindow object (IDispatch)
        """
        try:
            log.info("PowerPointEventSink: SlideShowBegin event received")
            if self._worker:
                self._worker.on_slideshow_begin(wn)
        except Exception as e:
            log.error(f"PowerPointEventSink: Error in SlideShowBegin - {e}")

    def SlideShowEnd(self, pres):
        """Called when slideshow ends.

        v0.0.56: Track slideshow state to modify announcements.

        Args:
            pres: Presentation object (IDispatch)
        """
        try:
            log.info("PowerPointEventSink: SlideShowEnd event received")
            if self._worker:
                self._worker.on_slideshow_end(pres)
        except Exception as e:
            log.error(f"PowerPointEventSink: Error in SlideShowEnd - {e}")

    def SlideShowNextSlide(self, slideShowWindow):
        """Called when slide advances in slideshow mode.

        Args:
            slideShowWindow: SlideShowWindow object (IDispatch)
        """
        try:
            log.debug("PowerPointEventSink: SlideShowNextSlide event received")
            if self._worker and slideShowWindow:
                try:
                    # Get slide index from slideshow window
                    slide_index = slideShowWindow.View.Slide.SlideIndex
                    if slide_index != self._last_slide_index:
                        log.info(f"PowerPointEventSink: Slideshow slide changed to {slide_index}")
                        self._last_slide_index = slide_index
                        # v0.0.56: Pass slideshow window for notes access
                        self._worker.on_slideshow_slide_changed(slide_index, slideShowWindow)
                except Exception as e:
                    log.debug(f"PowerPointEventSink: Could not get slideshow slide - {e}")
        except Exception as e:
            log.error(f"PowerPointEventSink: Error in SlideShowNextSlide - {e}")


# ============================================================================
# PowerPoint Worker Thread
# ============================================================================

class PowerPointWorker:
    """Background thread for PowerPoint COM operations.

    Runs COM work in a dedicated thread with proper STA initialization.
    Communicates UI updates back to main thread via queueHandler.

    v0.0.16: Uses COM events instead of polling for slide change detection.
    v0.0.17: Fixed import error in type library loading.
    v0.0.18: Multiple approaches to load type library (app object, GUID, registry).
    v0.0.19: Use wireEApplication from NVDA's built-in module instead of loading type library.
    v0.0.20: Access wireEApplication via direct module import (not relying on import *).
    v0.0.21: Define EApplication interface locally, use _AdviseConnection (NVDA pattern).
    v0.0.22: Use sel.Parent to get correct window when multiple presentations open.
    v0.0.23: Fix COM threading - queue navigation requests to worker thread.
    v0.0.24: Track Comments pane state to avoid redundant open/toggle calls.
    v0.0.25: Check actual pane visibility via GetPressedMso before calling ExecuteMso.
    v0.0.26: Remove session flag, rely only on GetPressedMso. Send F6 to focus Comments pane.
    v0.0.27: Add UIA diagnostic logging to find stable identifiers for Comments pane focus.
    v0.0.28: Replace F6 with parent chain walk to verify focus in Comments pane.
    v0.0.29: Add UIAutomationId diagnostic logging to find stable pane identifier.
    v0.0.30: Log UIAutomationId on every focus change via event_gainFocus.
    v0.0.31: Clean up diagnostic logging, use UIAutomationId for _is_in_comments_pane.
    v0.0.32: Add comment card diagnostic logging to analyze name/states for trimming.
    v0.0.33: Reformat comment announcements to "Author: comment" or "Resolved - Author: comment".
    v0.0.34: Fix comment detection - use name-based fallback when UIAAutomationId not available.
    v0.0.35: Add debug logging to diagnose why reformatting is not triggering.
    v0.0.36: Use event_NVDAObject_init for comment reformatting (NVDA recommended pattern).
    v0.0.37: Cancel-and-reannounce approach - speech.cancelSpeech() + ui.message() in event_gainFocus.
    v0.0.38: Add diagnostic logging to understand why event_gainFocus not firing for comments.
    v0.0.39: Log description property to find where comment text lives.
    v0.0.40: Debug logging to trace author/comment_text extraction.
    v0.0.41: Additional parse debug logging to find why author extraction fails.
    v0.0.42: Fix whitespace - normalize non-breaking spaces (U+00A0) from PowerPoint.
    v0.0.43: Also reformat reply comments (postRoot_) - strip date/time, announce as "Reply - Author: text".
    v0.0.44: Auto-tab from NewCommentButton to first comment on initial pane entry.
    v0.0.45: Fix auto-tab on PageUp/PageDown slide navigation - reset flag before navigate.
    v0.0.46: Use _pending_auto_focus flag for reliable auto-tab after slide navigation.
    v0.0.47: Announce slide number and title on PageUp/PageDown navigation.
    v0.0.48: Skip cancelSpeech after slide navigation to avoid cutting off title.
    v0.0.49: Add slide notes detection and Ctrl+Alt+N shortcut to read notes.
    v0.0.50: Fix double slide title announcement; strip **** and tag markers from notes.
    v0.0.51: Remove "Notes:" prefix; shorten "No notes on this slide" to "No notes".
    v0.0.52: Only detect/read notes with **** markers (meeting notes).
    v0.0.53: Extract only text BETWEEN **** markers, ignore text before/after.
    v0.0.54: Don't announce slide when NVDA starts with PowerPoint not focused.
    v0.0.55: Add detailed UIA logging for comment types (resolved, removed, status).
    v0.0.56: Slideshow mode - skip comment announcements, keep meeting notes; simplify reply/task status.
    v0.0.57: Debug logging for false "has meeting notes" during slideshow; fix premature SlideShowEnd.
    v0.0.58: Fix stuck slideshow state - use COM SlideShowWindows.Count instead of unreliable events.
    v0.0.59: Change "has meeting notes" to "has notes"; notes announced first in slideshow.
    v0.0.60: Fix first alt-tab announcement - don't mark slide as announced when focus check fails.
    v0.0.61: Slideshow notes via CustomSlideShowWindow._get_name() override - integrated single announcement.
    v0.0.62: Add reportFocus() override and extensive diagnostics - debug why custom class not instantiated.
    v0.0.63: Fix notes detection in slideshow - use self.currentSlide directly instead of worker thread.
    """

    # View type constants
    PP_VIEW_NORMAL = 9
    PP_VIEW_SLIDE_SORTER = 5
    PP_VIEW_NOTES = 10
    PP_VIEW_OUTLINE = 6
    PP_VIEW_SLIDE_MASTER = 3
    PP_VIEW_READING = 50

    def __init__(self):
        self._stop_event = threading.Event()
        self._thread = None
        self._ppt_app = None
        self._event_sink = None
        self._event_connection = None
        self._initialized = False
        # Track last slide for duplicate detection
        self._last_announced_slide = -1
        # v0.0.22: Store current window for correct multi-presentation support
        self._current_window = None
        # v0.0.23: Queue for navigation requests from main thread
        self._nav_request = None  # Will be direction: 1 for next, -1 for previous
        self._read_notes_request = False  # v0.0.49: Request to read slide notes
        self._from_comments_navigation = False  # v0.0.50: Track if nav from Comments pane
        self._has_received_focus = False  # v0.0.54: Track if app has received focus
        self._in_slideshow = False  # v0.0.56: Track if in presentation mode
        self._slideshow_window = None  # v0.0.56: Store slideshow window for notes access

    def start(self):
        """Start the background thread."""
        try:
            self._thread = threading.Thread(
                target=self._run,
                name="PowerPointCommentWorker",
                daemon=False  # Non-daemon for clean shutdown
            )
            self._thread.start()
            log.info("PowerPoint worker thread started")
        except Exception as e:
            log.error(f"Failed to start worker thread: {e}")

    def stop(self, timeout=5):
        """Stop the thread gracefully."""
        log.info("PowerPoint worker thread stopping...")
        self._stop_event.set()
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=timeout)
            if self._thread.is_alive():
                log.warning("PowerPoint worker thread did not stop within timeout")
            else:
                log.info("PowerPoint worker thread stopped cleanly")

    def request_initialize(self):
        """Request initialization from main thread.

        This is called from event_appModule_gainFocus.
        Sets a flag that the worker thread will pick up.
        v0.0.54: Also sets _has_received_focus to enable initial slide announcement.
        """
        log.info("Worker: Initialize requested (app has focus)")
        self._has_received_focus = True  # v0.0.54: App has focus, ok to announce
        self._initialized = False  # Force re-initialization

    def request_navigate(self, direction, from_comments_pane=False):
        """Request slide navigation from main thread.

        v0.0.23: This queues the request for the worker thread to process.
        COM objects can only be used on the thread that created them.
        v0.0.50: Added from_comments_pane flag to control slide title announcement.

        Args:
            direction: 1 for next slide, -1 for previous slide
            from_comments_pane: True if navigation triggered from Comments pane
        """
        log.info(f"Worker: Navigation requested (direction={direction}, from_comments={from_comments_pane})")
        self._nav_request = direction
        self._from_comments_navigation = from_comments_pane

    def request_read_notes(self):
        """Request to read slide notes from main thread.

        v0.0.49: Queues request for worker thread to read and announce notes.
        """
        log.info("Worker: Read notes requested")
        self._read_notes_request = True

    def _run(self):
        """Main thread loop - runs in background."""
        # Initialize COM in STA mode (required for Office)
        try:
            CoInitializeEx(COINIT_APARTMENTTHREADED)
            log.info("PowerPoint worker: COM initialized (STA)")
        except Exception as e:
            log.error(f"PowerPoint worker: Failed to initialize COM - {e}")
            return

        try:
            # Initial connection attempt
            self._initialize_com()

            # Main loop - process Windows messages to receive COM events
            while not self._stop_event.is_set():
                try:
                    # Pump Windows messages to receive COM events
                    # This is REQUIRED for COM events to be delivered
                    self._pump_messages(timeout_ms=500)

                    # Check if we need to reinitialize (e.g., after focus regained)
                    if not self._initialized:
                        self._initialize_com()

                    # v0.0.23: Check for navigation requests from main thread
                    if self._nav_request is not None:
                        direction = self._nav_request
                        self._nav_request = None  # Clear before processing
                        self._navigate_slide(direction)

                    # v0.0.49: Check for read notes requests from main thread
                    if self._read_notes_request:
                        self._read_notes_request = False
                        self._announce_slide_notes()

                except Exception as e:
                    log.error(f"Worker thread error: {e}")

        finally:
            # Clean up event connection
            self._disconnect_events()
            # Clean up COM
            self._ppt_app = None
            CoUninitialize()
            log.info("PowerPoint worker: COM uninitialized, thread exiting")

    def _pump_messages(self, timeout_ms=500):
        """Pump Windows messages to receive COM events.

        COM events are delivered via Windows messages, so we need to
        process the message queue for events to fire.

        Args:
            timeout_ms: How long to wait for messages (milliseconds)
        """
        try:
            # Use MsgWaitForMultipleObjects to wait for messages or timeout
            from ctypes import windll, byref, c_uint
            from ctypes.wintypes import MSG

            user32 = windll.user32

            # Wait for messages with timeout
            QS_ALLINPUT = 0x04FF
            WAIT_TIMEOUT = 0x00000102

            result = user32.MsgWaitForMultipleObjects(
                0,      # nCount - no handles
                None,   # pHandles
                False,  # bWaitAll
                timeout_ms,
                QS_ALLINPUT
            )

            # Process any pending messages
            msg = MSG()
            PM_REMOVE = 0x0001
            while user32.PeekMessageW(byref(msg), None, 0, 0, PM_REMOVE):
                user32.TranslateMessage(byref(msg))
                user32.DispatchMessageW(byref(msg))

        except Exception as e:
            log.debug(f"Message pump error (non-critical): {e}")

    def _initialize_com(self):
        """Connect to PowerPoint and set up event handling."""
        try:
            log.info("Worker: Attempting to connect to PowerPoint...")

            self._ppt_app = comHelper.getActiveObject(
                "PowerPoint.Application",
                dynamic=True
            )
            log.info("Worker: Connected to PowerPoint COM")

            if self._has_active_presentation():
                log.info("Worker: Active presentation found")
                self._ensure_normal_view()

                # Set up COM event handling
                self._connect_events()

                # Announce initial slide status
                self._check_initial_slide()

                self._initialized = True
            else:
                log.info("Worker: No active presentation - will retry on next focus")
                self._initialized = False

        except OSError as e:
            log.info(f"Worker: PowerPoint COM not available - {e}")
            self._ppt_app = None
            self._initialized = False
        except Exception as e:
            log.error(f"Worker: Initialize failed - {e}")
            self._ppt_app = None
            self._initialized = False

    def _connect_events(self):
        """Connect to PowerPoint application events.

        v0.0.21: Uses _AdviseConnection with locally-defined EApplication interface.
        This is the same pattern NVDA's built-in PowerPoint module uses.
        """
        if not self._ppt_app:
            log.warning("Worker: Cannot connect events - no PowerPoint app")
            return

        try:
            # First disconnect any existing connection
            self._disconnect_events()

            # Create event sink with our locally-defined EApplication interface
            self._event_sink = PowerPointEventSink(self)
            log.info("Worker: Created PowerPointEventSink")

            # Get IUnknown from sink for advise connection
            sink_iunknown = self._event_sink.QueryInterface(comtypes.IUnknown)
            log.info("Worker: Got IUnknown from sink")

            # Connect using _AdviseConnection (NOT GetEvents)
            # This is how NVDA's built-in powerpnt.py connects to events
            self._event_connection = _AdviseConnection(
                self._ppt_app,
                EApplication,
                sink_iunknown
            )

            log.info("Worker: Connected to PowerPoint events via _AdviseConnection")

        except Exception as e:
            log.error(f"Worker: Failed to connect to PowerPoint events - {e}")
            import traceback
            log.error(f"Worker: Traceback: {traceback.format_exc()}")
            self._event_sink = None
            self._event_connection = None

    def _disconnect_events(self):
        """Disconnect from PowerPoint events."""
        if self._event_connection:
            try:
                # The connection object should be released
                del self._event_connection
                self._event_connection = None
                log.info("Worker: Disconnected from PowerPoint events")
            except Exception as e:
                log.debug(f"Worker: Error disconnecting events - {e}")

        self._event_sink = None

    def _check_initial_slide(self):
        """Check and announce comments on the initial slide.

        v0.0.23: Skip if we already announced this slide (prevents double
        announcements on reinit after focus regained).
        v0.0.54: Only announce if app has received focus (prevents announcement
        when NVDA starts with PowerPoint open but not focused).
        v0.0.60: Don't mark slide as announced if we skip due to no focus.
        """
        try:
            current_index = self._get_current_slide_index()
            if current_index <= 0:
                return

            log.info(f"Worker: Initial slide is {current_index}, last announced was {self._last_announced_slide}")

            # v0.0.54: Don't announce if app hasn't received focus yet
            # v0.0.60: Also don't mark as announced, so we can announce on first real focus
            if not self._has_received_focus:
                log.info("Worker: Skipping initial slide announcement - app not focused yet (will announce on focus)")
                return

            # v0.0.23: Don't re-announce if same slide (prevents flashing on reinit)
            if current_index == self._last_announced_slide:
                log.info("Worker: Skipping re-announcement - same slide")
                return

            self._last_announced_slide = current_index
            self._announce_slide_comments()
        except Exception as e:
            log.debug(f"Worker: Error checking initial slide - {e}")

    def on_slide_changed_event(self, slide_index, window=None):
        """Called by event sink when slide changes.

        This runs on the COM thread (our worker thread).

        Args:
            slide_index: New slide index (1-based)
            window: The specific DocumentWindow that triggered the event (v0.0.22)
        """
        log.info(f"Worker: Slide change event received - slide {slide_index}")

        # v0.0.22: Store the window for use by other methods
        if window:
            self._current_window = window
            log.debug("Worker: Using specific window from event")
        elif not self._current_window and self._ppt_app:
            self._current_window = self._ppt_app.ActiveWindow
            log.debug("Worker: Falling back to ActiveWindow")

        # Avoid duplicate announcements
        if slide_index == self._last_announced_slide:
            log.debug(f"Worker: Ignoring duplicate slide {slide_index}")
            return

        self._last_announced_slide = slide_index

        # Ensure Normal view
        self._ensure_normal_view()

        # Announce comments on new slide
        self._announce_slide_comments()

    def on_slideshow_begin(self, wn):
        """Called when slideshow starts.

        v0.0.56: Track slideshow state to suppress comment announcements.

        Args:
            wn: SlideShowWindow object (IDispatch)
        """
        log.info("Worker: Slideshow started - entering presentation mode")
        self._in_slideshow = True
        self._slideshow_window = wn

    def on_slideshow_end(self, pres):
        """Called when slideshow ends.

        v0.0.56: Reset slideshow state.
        v0.0.57: Only reset if we were actually in slideshow mode (avoid false events).

        Args:
            pres: Presentation object (IDispatch)
        """
        if self._in_slideshow:
            log.info("Worker: Slideshow ended - exiting presentation mode")
            self._in_slideshow = False
            self._slideshow_window = None
        else:
            log.debug("Worker: SlideShowEnd received but not in slideshow mode - ignoring")

    def on_slideshow_slide_changed(self, slide_index, slideshow_window):
        """Called when slide changes during slideshow.

        v0.0.56: During slideshow, only announce meeting notes status (not comments).
        v0.0.61: Removed "has notes" announcement - now handled by CustomSlideShowWindow._get_name()

        The window name change will trigger NVDA's automatic announcement which will include
        "has notes, " prefix if present, providing a single integrated announcement.

        Args:
            slide_index: New slide index (1-based)
            slideshow_window: SlideShowWindow object for notes access
        """
        log.info(f"Worker: Slideshow slide changed to {slide_index}")

        # Store slideshow window for notes access (used by CustomSlideShowWindow._get_name())
        self._slideshow_window = slideshow_window

        # Avoid duplicate announcements
        if slide_index == self._last_announced_slide:
            log.debug(f"Worker: Ignoring duplicate slideshow slide {slide_index}")
            return

        self._last_announced_slide = slide_index

        # v0.0.61: Notes announcement now handled by CustomSlideShowWindow._get_name()
        # No additional announcement needed here - window name change handles it
        log.debug("Worker: Slideshow slide tracking updated (announcement via window name)")

    def _has_active_presentation(self):
        """Check if there's an active presentation open."""
        try:
            if self._ppt_app:
                if self._ppt_app.Presentations.Count > 0:
                    if self._ppt_app.ActiveWindow:
                        return True
            return False
        except Exception as e:
            log.debug(f"No active presentation: {e}")
            return False

    def _is_slideshow_running(self):
        """Check if a slideshow is actually running using COM.

        v0.0.58: Use SlideShowWindows.Count to verify slideshow state,
        rather than relying on potentially unreliable events.
        """
        try:
            if self._ppt_app:
                count = self._ppt_app.SlideShowWindows.Count
                log.debug(f"Worker: SlideShowWindows.Count = {count}")
                return count > 0
        except Exception as e:
            log.debug(f"Worker: Could not check slideshow state - {e}")
        return False

    def _get_window(self):
        """Get the current window (v0.0.22: prefer stored window over ActiveWindow)."""
        if self._current_window:
            return self._current_window
        if self._ppt_app:
            return self._ppt_app.ActiveWindow
        return None

    def _get_current_view(self):
        """Get current PowerPoint view type."""
        try:
            window = self._get_window()
            if window:
                view_type = window.ViewType
                log.debug(f"View type detected: {view_type}")
                return view_type
        except Exception as e:
            log.debug(f"Failed to get view type: {e}")
        return None

    def _ensure_normal_view(self):
        """Switch to Normal view if not already there."""
        try:
            window = self._get_window()
            current_view = self._get_current_view()
            if current_view is not None and current_view != self.PP_VIEW_NORMAL:
                log.info(f"Switching view from {current_view} to Normal")
                if window:
                    window.ViewType = self.PP_VIEW_NORMAL
                    self._announce("Switched to Normal view")
                    return True
            else:
                log.debug("Already in Normal view")
        except Exception as e:
            log.debug(f"Failed to switch view: {e}")
        return False

    def _announce(self, message):
        """Safely announce message on main thread."""
        try:
            log.info(f"Worker: Announcing '{message}'")
            queueFunction(eventQueue, ui.message, message)
        except Exception as e:
            log.error(f"Failed to queue announcement: {e}")

    def _get_current_slide_index(self):
        """Get current slide index (1-based)."""
        try:
            window = self._get_window()
            if window:
                return window.View.Slide.SlideIndex
        except Exception as e:
            log.debug(f"Worker: Could not get slide index - {e}")
        return -1

    def _get_slide_title(self):
        """Get the title text of the current slide.

        v0.0.47: Uses Shapes.HasTitle and Shapes.Title to get slide title.
        Returns empty string if slide has no title placeholder or title is empty.
        """
        try:
            window = self._get_window()
            if window:
                slide = window.View.Slide
                if slide.Shapes.HasTitle:
                    title_shape = slide.Shapes.Title
                    if title_shape and title_shape.HasTextFrame:
                        text_frame = title_shape.TextFrame
                        if text_frame.HasText:
                            return text_frame.TextRange.Text.strip()
        except Exception as e:
            log.debug(f"Worker: Could not get slide title - {e}")
        return ""

    def _get_slide_notes(self):
        """Get the notes text for the current slide.

        v0.0.49: Uses NotesPage.Shapes.Placeholders(2) to access notes text.
        Placeholder(1) is slide thumbnail, Placeholder(2) is notes body.
        v0.0.56: Uses SlideShowWindow when in presentation mode for proper sync.
        v0.0.57: Added debug logging for slide index to diagnose sync issues.

        Returns empty string if slide has no notes.

        References:
        - https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.notespage
        - https://learn.microsoft.com/en-us/office/vba/api/powerpoint.textframe
        """
        try:
            slide = None
            source = "unknown"

            # v0.0.56: Use SlideShowWindow when in presentation mode
            if self._in_slideshow and self._slideshow_window:
                try:
                    slide = self._slideshow_window.View.Slide
                    source = "SlideShowWindow"
                    log.debug(f"Worker: Getting notes from SlideShowWindow (slide {slide.SlideIndex})")
                except Exception as e:
                    log.debug(f"Worker: Could not get slide from SlideShowWindow - {e}")

            # Fall back to normal window
            if not slide:
                window = self._get_window()
                if window:
                    slide = window.View.Slide
                    source = "normal window"
                    log.debug(f"Worker: Getting notes from normal window (slide {slide.SlideIndex})")

            if slide:
                slide_idx = slide.SlideIndex
                notes_page = slide.NotesPage
                # Placeholder 2 is the notes body text
                placeholder = notes_page.Shapes.Placeholders(2)
                if placeholder.HasTextFrame:
                    text_frame = placeholder.TextFrame
                    if text_frame.HasText:
                        notes_text = text_frame.TextRange.Text.strip()
                        log.debug(f"Worker: Got notes for slide {slide_idx} from {source} ({len(notes_text)} chars)")
                        return notes_text
                log.debug(f"Worker: No notes text for slide {slide_idx} from {source}")
        except Exception as e:
            log.debug(f"Worker: Could not get slide notes - {e}")
        return ""

    def _has_meeting_notes(self):
        """Check if current slide has meeting notes (marked with ****).

        v0.0.49: Returns True if slide has non-empty notes text.
        v0.0.52: Only returns True if notes contain **** markers (meeting notes).
        Regular notes without markers are ignored.
        v0.0.57: Added debug logging to diagnose false positives.
        """
        notes = self._get_slide_notes()
        has_markers = '****' in notes if notes else False
        log.debug(f"Worker: _has_meeting_notes check - in_slideshow={self._in_slideshow}, "
                  f"notes_length={len(notes) if notes else 0}, has_markers={has_markers}")
        if notes:
            log.debug(f"Worker: Notes preview: {notes[:100]}...")
        if not notes:
            return False
        # Only consider notes with **** markers as "meeting notes"
        return has_markers

    def _clean_notes_text(self, notes):
        """Extract meeting notes content between **** markers.

        v0.0.50: Strips **** markers and <meeting notes> tags from notes.
        v0.0.53: Only extracts text BETWEEN **** markers, ignoring text before/after.
        """
        import re
        if not notes:
            return notes

        # v0.0.53: Extract only text between **** markers
        # Pattern: **** content **** (ignoring text before first and after last)
        marker_pattern = r'\*{4,}\s*(.*?)\s*\*{4,}'
        match = re.search(marker_pattern, notes, re.DOTALL)
        if match:
            cleaned = match.group(1)
        else:
            # Fallback: just remove markers (shouldn't happen if **** check passed)
            cleaned = re.sub(r'\*{4,}\s*', '', notes)

        # Remove <meeting notes> and </meeting notes> tags (case insensitive)
        cleaned = re.sub(r'</?meeting\s*notes>', '', cleaned, flags=re.IGNORECASE)
        # Remove <critical notes> and </critical notes> tags (case insensitive)
        cleaned = re.sub(r'</?critical\s*notes>', '', cleaned, flags=re.IGNORECASE)
        # Clean up extra whitespace
        cleaned = re.sub(r'\s+', ' ', cleaned).strip()
        return cleaned

    def _announce_slide_notes(self):
        """Announce the meeting notes text for current slide.

        v0.0.49: Called via Ctrl+Alt+N keyboard shortcut.
        v0.0.50: Strips marker tags before announcing.
        v0.0.51: Don't prefix with "Notes:" since user pressed the notes key.
        v0.0.52: Only reads notes that have **** markers (meeting notes).
        Regular notes without markers are ignored - says "No meeting notes".
        """
        notes = self._get_slide_notes()
        if notes and '****' in notes:
            cleaned = self._clean_notes_text(notes)
            self._announce(cleaned)
            log.info(f"Worker: Announced meeting notes ({len(cleaned)} chars)")
        else:
            self._announce("No notes")
            log.info("Worker: No notes on slide")

    def _get_comments_on_current_slide(self):
        """Get all comments on current slide."""
        try:
            window = self._get_window()
            if not window:
                log.debug("Worker: No window available for getting comments")
                return []

            slide = window.View.Slide
            comments = []
            comment_count = slide.Comments.Count
            log.debug(f"Worker: Found {comment_count} comments on slide")

            # COM collections are 1-indexed
            for i in range(1, comment_count + 1):
                try:
                    comment = slide.Comments.Item(i)
                    comments.append({
                        'text': comment.Text,
                        'author': comment.Author,
                        'datetime': comment.DateTime
                    })
                except Exception as e:
                    log.warning(f"Worker: Error reading comment {i} - {e}")

            return comments
        except Exception as e:
            log.debug(f"Worker: Could not get comments - {e}")
            return []

    def _announce_slide_comments(self):
        """Announce slide info and comment status for current slide.

        v0.0.47: Announces slide number and title first, then comment count.
        v0.0.49: Also announces "has notes" if slide has notes.
        v0.0.50: Only announce slide title if navigated from Comments pane (avoid double).
        v0.0.56: Skip comment announcements during slideshow (but keep meeting notes).
        v0.0.58: Use COM check instead of _in_slideshow flag (events unreliable).
        Format: "{slide_number}: {title}" then "Has X comments" or "No comments", then "has notes"
        """
        # v0.0.58: Check actual slideshow state via COM, not unreliable event flag
        # Events can fire out of order or get stuck
        actually_in_slideshow = self._is_slideshow_running()
        if actually_in_slideshow != self._in_slideshow:
            log.info(f"Worker: Fixing slideshow state mismatch - flag={self._in_slideshow}, actual={actually_in_slideshow}")
            self._in_slideshow = actually_in_slideshow
            if not actually_in_slideshow:
                self._slideshow_window = None

        if actually_in_slideshow:
            log.info("Worker: Skipping comment announcements - in slideshow mode")
            return

        # v0.0.50: Only announce slide title if navigated from Comments pane
        # NVDA's built-in PowerPoint module already announces slide on normal navigation
        if self._from_comments_navigation:
            slide_index = self._get_current_slide_index()
            slide_title = self._get_slide_title()

            if slide_index > 0:
                if slide_title:
                    slide_msg = f"{slide_index}: {slide_title}"
                else:
                    slide_msg = f"Slide {slide_index}"
                self._announce(slide_msg)
                log.info(f"Worker: {slide_msg}")
            # Reset flag after use
            self._from_comments_navigation = False

        # Announce comment count
        comments = self._get_comments_on_current_slide()

        if not comments:
            self._announce("No comments")
            log.info("Worker: No comments on this slide")
        else:
            count = len(comments)
            msg = f"Has {count} comment{'s' if count != 1 else ''}"
            self._announce(msg)
            log.info(f"Worker: {msg}")

            # Open Comments pane for slides with comments
            self._open_comments_pane()

        # v0.0.49: Announce if slide has meeting notes
        # v0.0.52: Only announce for notes with **** markers (meeting notes)
        if self._has_meeting_notes():
            self._announce("has notes")
            log.info("Worker: Slide has notes")

    def _is_comments_pane_visible(self):
        """Check if Comments pane is currently visible.

        v0.0.25: Uses GetMso to check actual command state.
        GetMso returns msoButtonDown (True/-1) if pane is open.

        Returns:
            True if Comments pane is visible, False otherwise
        """
        try:
            # Check the state of the CommentsPane toggle button
            # GetMso returns the pressed state: True/-1 if pressed (pane open)
            for cmd in ["CommentsPane", "ReviewShowComments", "ShowComments"]:
                try:
                    state = self._ppt_app.CommandBars.GetPressedMso(cmd)
                    log.info(f"Worker: GetPressedMso('{cmd}') = {state}")
                    if state:  # True or -1 means pressed/active
                        return True
                except Exception as e:
                    log.debug(f"Worker: GetPressedMso('{cmd}') failed - {e}")
                    continue
        except Exception as e:
            log.debug(f"Worker: Error checking pane visibility - {e}")
        return False

    def _open_comments_pane(self):
        """Open the Comments task pane if not already visible, then focus it.

        v0.0.25: Check actual pane visibility via GetPressedMso before toggling.
        v0.0.26: Removed session flag - rely only on GetPressedMso for accurate state.
                 Request F6 keypress to move focus to Comments pane.
        ExecuteMso toggles the pane, so calling it when already open would close it.
        """
        # v0.0.25: Check actual pane visibility
        if self._is_comments_pane_visible():
            log.info("Worker: Comments pane already visible - requesting focus")
            self._request_focus_comments_pane()
            return True

        # Pane is not visible - open it
        try:
            # CommentsPane is the correct idMso (confirmed working in v0.0.25)
            self._ppt_app.CommandBars.ExecuteMso("CommentsPane")
            log.info("Worker: Opened Comments pane via CommentsPane")
            # Request focus to move to the pane
            self._request_focus_comments_pane()
            return True
        except Exception as e:
            log.error(f"Worker: Error opening Comments pane - {e}")
        return False

    def _request_focus_comments_pane(self):
        """Request focus to Comments pane (placeholder for future implementation).

        v0.0.31: Diagnostic logging removed. Currently relies on PowerPoint's
        default behavior of focusing NewCommentButton when pane opens.
        Future: Could use UIA to find and focus first comment directly.
        """
        log.debug("Worker: Comments pane focus requested")

    def _navigate_slide(self, direction):
        """Navigate to next or previous slide (runs on worker thread).

        v0.0.22: Added for PageUp/PageDown support in Comments pane.
        v0.0.23: Renamed to private method - must run on worker thread.

        Args:
            direction: 1 for next slide, -1 for previous slide

        Returns:
            True if navigation succeeded, False otherwise
        """
        try:
            window = self._get_window()
            if not window:
                log.warning("Worker: No window for slide navigation")
                self._announce("Cannot navigate - no active presentation")
                return False

            current_index = window.View.Slide.SlideIndex
            presentation = window.Presentation
            total_slides = presentation.Slides.Count

            new_index = current_index + direction

            if new_index < 1:
                log.info("Worker: Already at first slide")
                self._announce("First slide")
                return False
            elif new_index > total_slides:
                log.info("Worker: Already at last slide")
                self._announce("Last slide")
                return False

            # Navigate to the new slide
            window.View.GotoSlide(new_index)
            log.info(f"Worker: Navigated to slide {new_index}")
            return True

        except Exception as e:
            log.error(f"Worker: Error navigating slide - {e}")
            self._announce("Navigation failed")
            return False


# ============================================================================
# Custom Overlay Classes
# ============================================================================

class CustomSlideShowWindow(SlideShowWindow):
    """Enhanced SlideShowWindow that announces speaker notes status in window name.

    v0.0.61: Initial implementation using _get_name() override (not working).
    v0.0.62: Added diagnostics and reportFocus() override to fix announcement path.

    Research found that NVDA slideshow announcements use handleSlideChange() → reportFocus(),
    not the normal focus event flow. The _get_name() override is correct but wasn't being
    called because reportFocus() needs to be overridden to ensure our custom class methods
    are invoked.

    This ensures notes status is announced BEFORE slide number/title as a single
    integrated announcement, not as a separate speech event.

    Example announcements:
    - With notes: "has notes, Slide show - Slide 3, Meeting Overview"
    - Without notes: "Slide show - Slide 3, Meeting Overview"
    - Notes mode: "has notes, Slide show notes - Slide 3, Meeting Overview"

    This approach is:
    - Non-timing-dependent (no race conditions)
    - Single announcement point (no duplication)
    - Respects NVDA's architecture (standard overlay class pattern)
    - Preserves all other NVDA features (verbosity, speech settings, etc.)
    """

    def __init__(self, *args, **kwargs):
        """Initialize custom slideshow window with diagnostic logging."""
        log.info("CustomSlideShowWindow.__init__() CALLED - Instance created")
        super().__init__(*args, **kwargs)

    def _check_slide_has_notes(self):
        """Check if current slide has meeting notes (marked with ****).

        v0.0.63: Use self.currentSlide directly instead of worker thread.
        This avoids threading issues and is more reliable.

        Returns:
            bool: True if slide has notes with **** markers
        """
        try:
            if not self.currentSlide:
                log.debug("CustomSlideShowWindow: No currentSlide available")
                return False

            # Access notes directly from the slide
            notes_page = self.currentSlide.NotesPage
            placeholder = notes_page.Shapes.Placeholders(2)

            if not placeholder.HasTextFrame:
                log.debug("CustomSlideShowWindow: No text frame in notes")
                return False

            text_frame = placeholder.TextFrame
            if not text_frame.HasText:
                log.debug("CustomSlideShowWindow: No text in notes frame")
                return False

            notes_text = text_frame.TextRange.Text.strip()
            has_markers = '****' in notes_text
            log.info(f"CustomSlideShowWindow: Slide {self.currentSlide.SlideIndex} notes check - "
                    f"length={len(notes_text)}, has_markers={has_markers}")

            return has_markers

        except Exception as e:
            log.error(f"CustomSlideShowWindow: Error checking notes - {e}")
            return False

    def reportFocus(self):
        """Override reportFocus to inject notes announcement.

        v0.0.62: Added this override because NVDA slideshow uses handleSlideChange()
        which calls reportFocus() directly. This is the actual entry point for
        slideshow announcements, not the normal focus event flow.

        v0.0.63: Use self.currentSlide directly to check notes instead of worker thread.
        """
        log.info("CustomSlideShowWindow.reportFocus() CALLED")

        # Check for notes using the slide object directly
        has_notes = self._check_slide_has_notes()
        log.info(f"CustomSlideShowWindow.reportFocus(): has_notes = {has_notes}")

        if has_notes:
            # Get the base name
            base_name = super()._get_name()
            log.info(f"CustomSlideShowWindow.reportFocus(): Announcing 'has notes, {base_name}'")
            # Speak custom sequence
            import ui
            ui.message(f"has notes, {base_name}")
        else:
            # Normal announcement
            log.debug("CustomSlideShowWindow.reportFocus(): No notes, using normal announcement")
            super().reportFocus()

    def _get_name(self):
        """Get window name with notes status prepended if present.

        v0.0.62: Added diagnostic logging to verify if this method is ever called.
        v0.0.63: Use self.currentSlide directly to check notes instead of worker thread.

        This property is queried by NVDA during focus reporting to determine
        what to announce. We prepend "has notes, " when the current slide
        has speaker notes with **** markers.

        Returns:
            str: Window name with optional "has notes, " prefix
        """
        log.info("CustomSlideShowWindow._get_name() CALLED")

        # Get base announcement from parent class
        # This will be "Slide show - {slideName}" or "Slide show notes - {slideName}"
        base_name = super()._get_name()
        log.info(f"CustomSlideShowWindow._get_name(): Base name = '{base_name}'")

        # Check if current slide has meeting notes using slide object directly
        if self._check_slide_has_notes():
            log.info("CustomSlideShowWindow._get_name(): Slide has notes - prepending to announcement")
            return f"has notes, {base_name}"
        else:
            log.debug("CustomSlideShowWindow._get_name(): No meeting notes on slide")
            return base_name


# ============================================================================
# AppModule - NVDA Integration
# ============================================================================

# Inherit from the just-imported AppModule (NVDA doc pattern)
# This preserves all built-in PowerPoint support while adding our features
class AppModule(AppModule):
    """Enhanced PowerPoint with comment navigation.

    Extends NVDA's built-in PowerPoint support using the pattern from
    NVDA Developer Guide and Joseph Lee's Office Desk addon.

    Uses COM events for instant slide change detection (v0.0.16).
    v0.0.22: PageUp/PageDown navigation in Comments pane.
    """

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._worker = None
        try:
            self._worker = PowerPointWorker()
            self._worker.start()
            log.info(f"PowerPoint Comments AppModule instantiated (v{ADDON_VERSION})")
        except Exception as e:
            log.error(f"PowerPoint Comments: Failed to create worker - {e}")

    def event_appModule_gainFocus(self):
        """Called when PowerPoint gains focus.

        IMPORTANT: This is an optional hook - parent class doesn't define it.
        Do NOT call super() here - it will fail with AttributeError.

        Requests initialization from the worker thread.
        """
        log.info("PowerPoint Comments: App gained focus - requesting initialization")
        if self._worker:
            self._worker.request_initialize()
        else:
            log.warning("PowerPoint Comments: Worker not available, skipping initialization")

    def chooseNVDAObjectOverlayClasses(self, obj, clsList):
        """Apply custom overlay classes for PowerPoint objects.

        v0.0.61: Registers CustomSlideShowWindow to override built-in slideshow window.
        v0.0.62: Added extensive diagnostics to debug why custom class not instantiated.

        This allows us to customize the window name announcement to include "has notes"
        status before the slide number/title.

        Called by NVDA for each object to determine which custom classes to apply.
        We let the parent handle standard selection first, then replace SlideShowWindow
        with our CustomSlideShowWindow to get "has notes" announcement integration.

        Args:
            obj: The NVDAObject being initialized
            clsList: List of classes to apply (modified in place)
        """
        # Let parent class do its selection first
        # This populates clsList with built-in PowerPoint classes
        super().chooseNVDAObjectOverlayClasses(obj, clsList)

        # Check if parent assigned SlideShowWindow for this object
        # SlideShowWindow is only used for slideshow presentation windows
        if SlideShowWindow in clsList:
            try:
                # v0.0.62: Log object details when SlideShowWindow is in the list
                obj_info = f"windowClassName={getattr(obj, 'windowClassName', 'N/A')}, role={getattr(obj, 'role', 'N/A')}"
                slideshow_state = getattr(self._worker, '_is_slideshow', None) if self._worker else None
                log.info(f"chooseNVDAObjectOverlayClasses: SlideShowWindow found! {obj_info}, _is_slideshow={slideshow_state}")
                log.info(f"chooseNVDAObjectOverlayClasses: BEFORE replacement - clsList={[cls.__name__ for cls in clsList]}")

                # Replace built-in SlideShowWindow with our custom version
                idx = clsList.index(SlideShowWindow)
                clsList[idx] = CustomSlideShowWindow
                log.info(f"chooseNVDAObjectOverlayClasses: REPLACED at index {idx}")
                log.info(f"chooseNVDAObjectOverlayClasses: AFTER replacement - clsList={[cls.__name__ for cls in clsList]}")
            except Exception as e:
                log.error(f"chooseNVDAObjectOverlayClasses: Error replacing SlideShowWindow - {e}")

    def event_gainFocus(self, obj, nextHandler):
        """Called when any object in PowerPoint gains focus.

        v0.0.37: Cancel-and-reannounce approach for comment cards.
        v0.0.38: Re-added try/except - was causing silent crashes.
        v0.0.43: Also reformat reply comments (postRoot_) to remove date/time.
        v0.0.44: Auto-tab from NewCommentButton to first comment.
        v0.0.46: Use _pending_auto_focus for reliable auto-tab after slide navigation.
        """
        try:
            import re
            from inputCore import manager as inputManager
            from keyboardHandler import KeyboardInputGesture

            uia_id = getattr(obj, 'UIAAutomationId', '') or ''
            name = getattr(obj, 'name', '') or ''
            description = getattr(obj, 'description', '') or ''
            role = getattr(obj, 'role', None)
            role_name = getattr(obj, 'roleText', '') or ''
            states = getattr(obj, 'states', set()) or set()

            # Normalize whitespace - PowerPoint uses non-breaking spaces (U+00A0)
            name_normalized = re.sub(r'\s+', ' ', name)

            # v0.0.55: Detailed UIA logging for comment types (resolved, removed, status changes)
            # Log all comment-related elements for research purposes
            is_comment_element = (
                uia_id.startswith('cardRoot_') or
                uia_id.startswith('postRoot_') or
                uia_id == 'NewCommentButton' or
                uia_id == 'CommentsList' or
                'comment' in uia_id.lower() or
                'comment' in name.lower() or
                'resolved' in name.lower() or
                'removed' in name.lower()
            )
            if is_comment_element:
                log.info(f"=== UIA COMMENT ELEMENT ===")
                log.info(f"  UIAAutomationId: {uia_id}")
                log.info(f"  Name: {name[:200] if name else '(empty)'}")
                log.info(f"  Name normalized: {name_normalized[:200] if name_normalized else '(empty)'}")
                log.info(f"  Description: {description[:200] if description else '(empty)'}")
                log.info(f"  Role: {role} ({role_name})")
                log.info(f"  States: {states}")
                # Log additional UIA properties if available
                try:
                    class_name = getattr(obj, 'UIAClassName', '') or ''
                    log.info(f"  ClassName: {class_name}")
                except:
                    pass
                try:
                    control_type = getattr(obj, 'UIAControlType', '') or ''
                    log.info(f"  ControlType: {control_type}")
                except:
                    pass
                log.info(f"=== END UIA COMMENT ELEMENT ===")

            # Detect if we're in the Comments pane (NewCommentButton, cardRoot_, or postRoot_)
            is_in_comments = (
                uia_id == 'NewCommentButton' or
                uia_id.startswith('cardRoot_') or
                uia_id.startswith('postRoot_')
            )

            # Reset flags when leaving Comments pane
            if not is_in_comments:
                self._in_comments_pane = False
                self._pending_auto_focus = False

            # v0.0.46: Check for pending auto-focus (set by PageUp/PageDown navigation)
            # This triggers when ANY comments pane element gets focus after slide change
            # v0.0.48: Set _just_navigated flag to skip cancelSpeech for first comment
            if is_in_comments and getattr(self, '_pending_auto_focus', False):
                self._pending_auto_focus = False
                self._in_comments_pane = True
                self._just_navigated = True  # v0.0.48: Don't cancel speech for this comment
                # If we landed on NewCommentButton, tab to first comment
                if uia_id == 'NewCommentButton':
                    log.info("Auto-focus after slide change - tabbing to first comment")
                    KeyboardInputGesture.fromName("tab").send()
                    return  # Don't announce the button
                # If we landed directly on a comment, just mark as in pane (no tab needed)
                else:
                    log.info(f"Auto-focus after slide change - already on comment: {uia_id[:30]}")

            # v0.0.44: Auto-tab from NewCommentButton on initial F6 entry
            elif uia_id == 'NewCommentButton':
                # Only auto-tab if we're entering the Comments pane for the first time
                # (not if user navigated back to the button)
                if not getattr(self, '_in_comments_pane', False):
                    self._in_comments_pane = True
                    log.info("Entering Comments pane via F6 - auto-tabbing to first comment")
                    # Send Tab key to move to first comment
                    KeyboardInputGesture.fromName("tab").send()
                    return  # Don't announce the button

            # Check if this is a comment thread card (cardRoot_)
            is_comment_card = (
                uia_id.startswith('cardRoot_') or
                'Comment thread started by' in name_normalized
            )

            # Check if this is a reply comment (postRoot_)
            is_reply_comment = (
                uia_id.startswith('postRoot_') or
                name_normalized.startswith('Comment by ')
            )

            if is_comment_card:
                # Extract author and resolved state for thread cards
                is_resolved = name_normalized.startswith("Resolved ")
                author = ""

                if " started by " in name_normalized:
                    author_part = name_normalized.split(" started by ", 1)[1]
                    if ", with " in author_part:
                        author = author_part.split(", with ")[0]
                    else:
                        author = author_part

                if author and description:
                    # v0.0.48: Skip cancelSpeech after slide navigation to let title finish
                    if not getattr(self, '_just_navigated', False):
                        speech.cancelSpeech()
                    else:
                        self._just_navigated = False  # Clear flag after use
                        log.info("Skipped cancelSpeech - letting slide title finish")
                    if is_resolved:
                        formatted = f"Resolved - {author}: {description}"
                    else:
                        formatted = f"{author}: {description}"
                    ui.message(formatted)
                    log.info(f"Comment reformatted: {formatted[:80]}")
                    return

            elif is_reply_comment:
                # v0.0.56: Handle multiple reply/status formats:
                # - "Comment by Author on Month Day, Year, Time" -> "Author: description"
                # - "Task updated by Author on Month Day, Year, Time" -> "Author - description"
                author = ""
                is_task_status = False

                if name_normalized.startswith("Task updated by "):
                    # Task status update: "Task updated by Author on ..."
                    # Description will be "Completed a task" or "Reopened a task"
                    after_prefix = name_normalized[16:]  # Skip "Task updated by "
                    if " on " in after_prefix:
                        author = after_prefix.split(" on ", 1)[0]
                    is_task_status = True
                elif name_normalized.startswith("Comment by "):
                    # Regular reply: "Comment by Author on ..."
                    after_prefix = name_normalized[11:]  # Skip "Comment by "
                    if " on " in after_prefix:
                        author = after_prefix.split(" on ", 1)[0]

                if author and description:
                    # v0.0.48: Skip cancelSpeech after slide navigation to let title finish
                    if not getattr(self, '_just_navigated', False):
                        speech.cancelSpeech()
                    else:
                        self._just_navigated = False  # Clear flag after use
                        log.info("Skipped cancelSpeech - letting slide title finish")

                    # v0.0.56: Format based on type
                    if is_task_status:
                        # Task status: "Author - Completed task" or "Author - Reopened task"
                        status_text = description.replace(" a task", " task")
                        formatted = f"{author} - {status_text}"
                    else:
                        # Regular reply: just "Author: description" (no "Reply -" prefix)
                        formatted = f"{author}: {description}"

                    ui.message(formatted)
                    log.info(f"Reply/Status reformatted: {formatted[:80]}")
                    return

        except Exception as e:
            log.error(f"event_gainFocus error: {e}")

        nextHandler()

    def _is_in_comments_pane(self):
        """Check if focus is currently in the Comments pane.

        v0.0.31: Uses UIAutomationId for reliable detection without localized text.
        Stable identifiers found in v0.0.30 testing:
        - NewCommentButton (exact)
        - CommentsList (exact)
        - cardRoot_ prefix (comment threads)
        - firstPaneElement prefix (pane container)
        """
        try:
            focus = api.getFocusObject()
            if not focus:
                return False

            # Check focused element and walk up parent chain
            obj = focus
            for _ in range(15):
                if obj is None:
                    break
                try:
                    # Get UIAutomationId - this is the stable identifier
                    uia_id = getattr(obj, 'UIAAutomationId', '') or ''

                    # Check for any Comments pane identifier
                    if (uia_id == 'NewCommentButton' or
                        uia_id == 'CommentsList' or
                        uia_id.startswith('cardRoot_') or
                        uia_id.startswith('firstPaneElement')):
                        log.debug(f"_is_in_comments_pane: MATCH - UIAutomationId='{uia_id}'")
                        return True
                except Exception:
                    pass
                obj = getattr(obj, 'parent', None)

        except Exception as e:
            log.error(f"_is_in_comments_pane: Error - {e}")
        return False

    @script(
        gesture="kb:pageDown",
        description="Navigate to next slide (in Comments pane)",
        category="PowerPoint Comments"
    )
    def script_nextSlideFromComments(self, gesture):
        """Navigate to next slide when in Comments pane.

        v0.0.22: PageDown switches slides while in Comments pane.
        v0.0.23: Use request_navigate() to queue for worker thread.
        v0.0.46: Set _pending_auto_focus to trigger auto-tab when focus returns.
        Otherwise, passes the key through to PowerPoint.
        """
        if self._is_in_comments_pane() and self._worker:
            log.info("PageDown in Comments pane - requesting next slide")
            # v0.0.46: Set pending flag so auto-tab triggers when focus returns
            self._pending_auto_focus = True
            self._in_comments_pane = False
            self._worker.request_navigate(1, from_comments_pane=True)
        else:
            # Pass through to PowerPoint
            gesture.send()

    @script(
        gesture="kb:pageUp",
        description="Navigate to previous slide (in Comments pane)",
        category="PowerPoint Comments"
    )
    def script_previousSlideFromComments(self, gesture):
        """Navigate to previous slide when in Comments pane.

        v0.0.22: PageUp switches slides while in Comments pane.
        v0.0.23: Use request_navigate() to queue for worker thread.
        v0.0.46: Set _pending_auto_focus to trigger auto-tab when focus returns.
        Otherwise, passes the key through to PowerPoint.
        """
        if self._is_in_comments_pane() and self._worker:
            log.info("PageUp in Comments pane - requesting previous slide")
            # v0.0.46: Set pending flag so auto-tab triggers when focus returns
            self._pending_auto_focus = True
            self._in_comments_pane = False
            self._worker.request_navigate(-1, from_comments_pane=True)
        else:
            # Pass through to PowerPoint
            gesture.send()

    @script(
        gesture="kb:control+alt+n",
        description="Read slide notes",
        category="PowerPoint Comments"
    )
    def script_readSlideNotes(self, gesture):
        """Read the notes for the current slide.

        v0.0.49: Ctrl+Alt+N announces slide notes.
        Works anywhere in PowerPoint (not just Comments pane).
        """
        if self._worker:
            log.info("Ctrl+Alt+N pressed - requesting notes read")
            self._worker.request_read_notes()
        else:
            ui.message("Notes not available")

    def terminate(self):
        """Clean up when PowerPoint closes or NVDA exits."""
        log.info("PowerPoint Comments: Terminating - stopping worker thread")
        if hasattr(self, '_worker') and self._worker:
            self._worker.stop(timeout=5)
        super().terminate()
