# appModules/powerpnt.py
# PowerPoint Comments NVDA Addon - Phase 1: Foundation + View Management
#
# This module extends NVDA's built-in PowerPoint support to add
# comment navigation features.
#
# Pattern: NVDA Developer Guide - extending built-in appModules
# https://download.nvaccess.org/documentation/developerGuide.html
# Uses: from nvdaBuiltin.appModules.xxx import * then class AppModule(AppModule)

# Addon version - update this and manifest.ini together
ADDON_VERSION = "0.0.14"

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
import queue
from comtypes import CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED
from queueHandler import queueFunction, eventQueue


class PowerPointWorker:
    """Background thread for PowerPoint COM operations.

    Runs COM work in a dedicated thread with proper STA initialization.
    Communicates UI updates back to main thread via queueHandler.
    """

    # View type constants (shared with AppModule)
    PP_VIEW_NORMAL = 9
    PP_VIEW_SLIDE_SORTER = 5
    PP_VIEW_NOTES = 10
    PP_VIEW_OUTLINE = 6
    PP_VIEW_SLIDE_MASTER = 3
    PP_VIEW_READING = 50

    def __init__(self):
        self._stop_event = threading.Event()
        self._work_queue = queue.Queue()
        self._thread = None
        self._ppt_app = None

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

    def queue_task(self, task_name, *args):
        """Queue a task for the background thread."""
        self._work_queue.put((task_name, args))
        log.debug(f"Queued task: {task_name}")

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
            while not self._stop_event.is_set():
                try:
                    # Check for work with timeout (allows stop check)
                    task_name, args = self._work_queue.get(timeout=0.5)
                    self._execute_task(task_name, args)
                except queue.Empty:
                    # No work, continue loop
                    pass
                except Exception as e:
                    log.error(f"Worker thread error: {e}")
        finally:
            # Always clean up COM
            self._ppt_app = None
            CoUninitialize()
            log.info("PowerPoint worker: COM uninitialized, thread exiting")

    def _execute_task(self, task_name, args):
        """Execute a queued task."""
        log.info(f"Worker executing task: {task_name}")
        if task_name == "initialize":
            self._task_initialize()
        # Add more tasks as needed for Phase 2

    def _task_initialize(self):
        """Connect to PowerPoint and check presentation."""
        try:
            self._ppt_app = comHelper.getActiveObject(
                "PowerPoint.Application",
                dynamic=True
            )
            log.info("Worker: Connected to PowerPoint COM")

            if self._has_active_presentation():
                log.info("Worker: Active presentation found")
                self._ensure_normal_view()
            else:
                log.info("Worker: No active presentation")
        except OSError as e:
            log.info(f"Worker: COM not ready - {e}")
            self._ppt_app = None
        except Exception as e:
            log.error(f"Worker: Initialize failed - {e}")
            self._ppt_app = None

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
                # Announce on main thread
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
            queueFunction(eventQueue, ui.message, message)
        except Exception as e:
            log.error(f"Failed to queue announcement: {e}")


# Inherit from the just-imported AppModule (NVDA doc pattern)
# This preserves all built-in PowerPoint support while adding our features
class AppModule(AppModule):
    """Enhanced PowerPoint with comment navigation.

    Extends NVDA's built-in PowerPoint support using the pattern from
    NVDA Developer Guide and Joseph Lee's Office Desk addon.

    Uses a dedicated background thread for all COM operations.
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

        Queues initialization task to background thread.
        """
        log.info("PowerPoint Comments: App gained focus - queuing initialization")
        if self._worker:
            self._worker.queue_task("initialize")
        else:
            log.warning("PowerPoint Comments: Worker not available, skipping initialization")

    def terminate(self):
        """Clean up when PowerPoint closes or NVDA exits."""
        log.info("PowerPoint Comments: Terminating - stopping worker thread")
        if hasattr(self, '_worker') and self._worker:
            self._worker.stop(timeout=5)
        super().terminate()
