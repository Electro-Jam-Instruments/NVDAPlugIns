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
ADDON_VERSION = "0.0.36"

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
                        self._worker.on_slide_changed_event(slide_index)
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
        """
        log.info("Worker: Initialize requested")
        self._initialized = False  # Force re-initialization

    def request_navigate(self, direction):
        """Request slide navigation from main thread.

        v0.0.23: This queues the request for the worker thread to process.
        COM objects can only be used on the thread that created them.

        Args:
            direction: 1 for next slide, -1 for previous slide
        """
        log.info(f"Worker: Navigation requested (direction={direction})")
        self._nav_request = direction

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
        """
        try:
            current_index = self._get_current_slide_index()
            if current_index > 0:
                log.info(f"Worker: Initial slide is {current_index}, last announced was {self._last_announced_slide}")
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
        """Announce comment status for current slide."""
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

    def event_NVDAObject_init(self, obj):
        """Called when an NVDA object is initialized - BEFORE properties are cached.

        v0.0.36: Use this instead of event_gainFocus to modify name before announcement.
        Per NVDA Developer Guide: this is the recommended approach for property modifications.
        https://download.nvaccess.org/documentation/developerGuide.html
        """
        try:
            uia_id = getattr(obj, 'UIAAutomationId', '') or ''
            name = getattr(obj, 'name', '') or ''

            # Check if this is a comment card
            is_comment_card = (
                uia_id.startswith('cardRoot_') or
                'Comment thread started by' in name
            )

            if is_comment_card:
                description = getattr(obj, 'description', '') or ''

                # Extract author and resolved state from name
                # Name format: "Comment thread started by Author" or
                #              "Resolved comment thread started by Author"
                is_resolved = name.startswith("Resolved ")
                author = ""

                if " started by " in name:
                    # Author may have suffix like ", with 1 reply"
                    author_part = name.split(" started by ", 1)[1]
                    # Remove suffix like ", with 1 reply"
                    if ", with " in author_part:
                        author = author_part.split(", with ")[0]
                    else:
                        author = author_part

                # Build and set new name
                if author and description:
                    if is_resolved:
                        obj.name = f"Resolved - {author}: {description}"
                    else:
                        obj.name = f"{author}: {description}"
                    log.info(f"Comment reformatted: {obj.name}")

        except Exception as e:
            log.debug(f"event_NVDAObject_init error: {e}")

    def event_gainFocus(self, obj, nextHandler):
        """Called when any object in PowerPoint gains focus.

        v0.0.36: Comment reformatting moved to event_NVDAObject_init.
        This handler kept for future use if needed.
        """
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
        Otherwise, passes the key through to PowerPoint.
        """
        if self._is_in_comments_pane() and self._worker:
            log.info("PageDown in Comments pane - requesting next slide")
            self._worker.request_navigate(1)
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
        Otherwise, passes the key through to PowerPoint.
        """
        if self._is_in_comments_pane() and self._worker:
            log.info("PageUp in Comments pane - requesting previous slide")
            self._worker.request_navigate(-1)
        else:
            # Pass through to PowerPoint
            gesture.send()

    def terminate(self):
        """Clean up when PowerPoint closes or NVDA exits."""
        log.info("PowerPoint Comments: Terminating - stopping worker thread")
        if hasattr(self, '_worker') and self._worker:
            self._worker.stop(timeout=5)
        super().terminate()
