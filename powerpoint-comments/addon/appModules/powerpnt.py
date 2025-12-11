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
ADDON_VERSION = "0.0.17"

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
import threading
import ctypes
from comtypes import CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED, COMObject
from comtypes.client import GetEvents
from queueHandler import queueFunction, eventQueue

# ============================================================================
# COM Event Sink for PowerPoint Application Events
# ============================================================================

# We need to get the EApplication interface from PowerPoint's type library
# This will be populated when we connect to PowerPoint
_ppt_events_interface = None


def _get_ppt_events_interface(ppt_app):
    """Get the EApplication events interface from PowerPoint type library.

    Returns the interface class or None if not available.
    Logs detailed error information on failure.
    """
    global _ppt_events_interface

    if _ppt_events_interface is not None:
        return _ppt_events_interface

    try:
        import comtypes.client

        log.info("Attempting to get PowerPoint events interface from type library...")

        # Generate the PowerPoint type library wrapper
        # This creates Python classes for all PowerPoint COM interfaces
        try:
            ppt_gen = comtypes.client.GetModule(['{91493440-5A91-11CF-8700-00AA0060263B}', 1, 0])
            log.info(f"PowerPoint type library loaded: {ppt_gen}")

            # Look for EApplication interface (the events interface)
            if hasattr(ppt_gen, 'EApplication'):
                _ppt_events_interface = ppt_gen.EApplication
                log.info(f"Found EApplication interface: {_ppt_events_interface}")
                return _ppt_events_interface
            else:
                log.warning("EApplication interface not found in type library")
                # List available interfaces for debugging
                interfaces = [name for name in dir(ppt_gen) if not name.startswith('_')]
                log.debug(f"Available interfaces: {interfaces[:20]}...")  # First 20

        except Exception as e:
            log.error(f"Failed to load PowerPoint type library: {e}")

    except Exception as e:
        log.error(f"Failed to get events interface: {e}")

    return None


class PowerPointEventSink(COMObject):
    """COM Event Sink for PowerPoint application events.

    Receives SlideSelectionChanged and WindowSelectionChange events from PowerPoint.
    Calls back to the worker thread to process slide changes.

    The _com_interfaces_ attribute is set dynamically when we connect,
    because we need the PowerPoint type library to be loaded first.
    """

    # Will be set dynamically when interface is discovered
    _com_interfaces_ = []

    def __init__(self, worker):
        """Initialize the event sink.

        Args:
            worker: PowerPointWorker instance to notify on events
        """
        super().__init__()
        self._worker = worker
        self._last_slide_index = -1
        log.info("PowerPointEventSink: Initialized")

    def SlideSelectionChanged(self, SldRange):
        """Called when slide selection changes in thumbnail pane.

        Args:
            SldRange: SlideRange object containing selected slides
        """
        try:
            log.debug("PowerPointEventSink: SlideSelectionChanged event received")
            if SldRange and SldRange.Count > 0:
                slide_index = SldRange.Item(1).SlideIndex
                log.info(f"PowerPointEventSink: Slide selected - index {slide_index}")

                if slide_index != self._last_slide_index:
                    self._last_slide_index = slide_index
                    # Notify worker of slide change
                    if self._worker:
                        self._worker.on_slide_changed_event(slide_index)
            else:
                log.debug("PowerPointEventSink: SlideSelectionChanged - no slides in range")
        except Exception as e:
            log.error(f"PowerPointEventSink: Error in SlideSelectionChanged - {e}")

    def WindowSelectionChange(self, Sel):
        """Called when selection changes in PowerPoint window.

        This is a broader event that fires for text, shape, and slide selections.
        We use it as a backup in case SlideSelectionChanged doesn't fire.

        Args:
            Sel: Selection object with Type property:
                 0=ppSelectionNone, 1=ppSelectionSlides, 2=ppSelectionShapes, 3=ppSelectionText
        """
        try:
            log.debug(f"PowerPointEventSink: WindowSelectionChange event received")
            # We could check Sel.Type here, but for now just check slide index
            if self._worker and self._worker._ppt_app:
                try:
                    current_index = self._worker._ppt_app.ActiveWindow.View.Slide.SlideIndex
                    if current_index != self._last_slide_index:
                        log.info(f"PowerPointEventSink: Slide change detected via WindowSelectionChange - {current_index}")
                        self._last_slide_index = current_index
                        self._worker.on_slide_changed_event(current_index)
                except Exception as e:
                    log.debug(f"PowerPointEventSink: Could not get slide in WindowSelectionChange - {e}")
        except Exception as e:
            log.error(f"PowerPointEventSink: Error in WindowSelectionChange - {e}")


# ============================================================================
# PowerPoint Worker Thread
# ============================================================================

class PowerPointWorker:
    """Background thread for PowerPoint COM operations.

    Runs COM work in a dedicated thread with proper STA initialization.
    Communicates UI updates back to main thread via queueHandler.

    v0.0.16: Uses COM events instead of polling for slide change detection.
    v0.0.17: Fixed import error in type library loading.
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
        """Connect to PowerPoint application events."""
        if not self._ppt_app:
            log.warning("Worker: Cannot connect events - no PowerPoint app")
            return

        try:
            # First disconnect any existing connection
            self._disconnect_events()

            # Get the events interface
            events_interface = _get_ppt_events_interface(self._ppt_app)

            if events_interface:
                # Set the interface on our sink class
                PowerPointEventSink._com_interfaces_ = [events_interface]

                # Create event sink
                self._event_sink = PowerPointEventSink(self)

                # Connect to events
                self._event_connection = GetEvents(self._ppt_app, self._event_sink)

                log.info("Worker: Connected to PowerPoint events successfully")
            else:
                log.error("Worker: Could not get PowerPoint events interface - slide change detection disabled")

        except Exception as e:
            log.error(f"Worker: Failed to connect to PowerPoint events - {e}")
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
        """Check and announce comments on the initial slide."""
        try:
            current_index = self._get_current_slide_index()
            if current_index > 0:
                log.info(f"Worker: Initial slide is {current_index}")
                self._last_announced_slide = current_index
                self._announce_slide_comments()
        except Exception as e:
            log.debug(f"Worker: Error checking initial slide - {e}")

    def on_slide_changed_event(self, slide_index):
        """Called by event sink when slide changes.

        This runs on the COM thread (our worker thread).

        Args:
            slide_index: New slide index (1-based)
        """
        log.info(f"Worker: Slide change event received - slide {slide_index}")

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

    def _get_current_view(self):
        """Get current PowerPoint view type."""
        try:
            if self._ppt_app and self._ppt_app.ActiveWindow:
                view_type = self._ppt_app.ActiveWindow.ViewType
                log.debug(f"View type detected: {view_type}")
                return view_type
        except Exception as e:
            log.debug(f"Failed to get view type: {e}")
        return None

    def _ensure_normal_view(self):
        """Switch to Normal view if not already there."""
        try:
            current_view = self._get_current_view()
            if current_view is not None and current_view != self.PP_VIEW_NORMAL:
                log.info(f"Switching view from {current_view} to Normal")
                self._ppt_app.ActiveWindow.ViewType = self.PP_VIEW_NORMAL
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
            if self._ppt_app and self._ppt_app.ActiveWindow:
                return self._ppt_app.ActiveWindow.View.Slide.SlideIndex
        except Exception as e:
            log.debug(f"Worker: Could not get slide index - {e}")
        return -1

    def _get_comments_on_current_slide(self):
        """Get all comments on current slide."""
        try:
            slide = self._ppt_app.ActiveWindow.View.Slide
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

    def _open_comments_pane(self):
        """Open the Comments task pane if not visible."""
        try:
            # Try multiple command names (varies by Office version)
            for cmd in ["ReviewShowComments", "ShowComments", "CommentsPane"]:
                try:
                    self._ppt_app.CommandBars.ExecuteMso(cmd)
                    log.info(f"Worker: Opened Comments pane via {cmd}")
                    return True
                except Exception as e:
                    log.debug(f"Worker: Command {cmd} failed - {e}")
                    continue
            log.warning("Worker: Could not open Comments pane - all commands failed")
        except Exception as e:
            log.error(f"Worker: Error opening Comments pane - {e}")
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

    def terminate(self):
        """Clean up when PowerPoint closes or NVDA exits."""
        log.info("PowerPoint Comments: Terminating - stopping worker thread")
        if hasattr(self, '_worker') and self._worker:
            self._worker.stop(timeout=5)
        super().terminate()
