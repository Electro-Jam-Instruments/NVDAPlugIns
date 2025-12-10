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
ADDON_VERSION = "0.0.9"

# Import logging FIRST so we can log any import issues
import logging
log = logging.getLogger(__name__)
log.info(f"PowerPoint Comments addon: Module loading (v{ADDON_VERSION})")

# Import EVERYTHING from built-in PowerPoint module
# This is the NVDA-documented pattern for extending built-in appModules
from nvdaBuiltin.appModules.powerpnt import *
log.info("PowerPoint Comments addon: Built-in powerpnt imported successfully")

# Additional imports for our functionality
from comtypes.client import GetActiveObject
import ui


# Inherit from the just-imported AppModule (NVDA doc pattern)
# This preserves all built-in PowerPoint support while adding our features
class AppModule(AppModule):
    """Enhanced PowerPoint with comment navigation.

    Extends NVDA's built-in PowerPoint support using the pattern from
    NVDA Developer Guide and Joseph Lee's Office Desk addon.
    """

    # View type constants
    PP_VIEW_NORMAL = 9
    PP_VIEW_SLIDE_SORTER = 5
    PP_VIEW_NOTES = 10
    PP_VIEW_OUTLINE = 6
    PP_VIEW_SLIDE_MASTER = 3
    PP_VIEW_READING = 50

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._ppt_app = None
        self._last_slide_index = -1
        log.info(f"PowerPoint Comments AppModule instantiated (v{ADDON_VERSION})")

    def event_appModule_gainFocus(self):
        """Called when PowerPoint gains focus."""
        log.info("PowerPoint Comments: App gained focus")
        if self._connect_to_powerpoint():
            if self._has_active_presentation():
                self._ensure_normal_view()
            else:
                log.debug("No active presentation - skipping view management")

    def _connect_to_powerpoint(self):
        """Connect to running PowerPoint instance."""
        try:
            self._ppt_app = GetActiveObject("PowerPoint.Application")
            log.info("PowerPoint Comments: Connected to COM")
            return True
        except OSError as e:
            # WinError -2147221021: Operation unavailable
            # This happens when PowerPoint is starting up or COM isn't ready
            log.debug(f"PowerPoint COM not ready: {e}")
            self._ppt_app = None
            return False
        except Exception as e:
            log.error(f"Failed to connect to PowerPoint: {e}")
            self._ppt_app = None
            return False

    def _has_active_presentation(self):
        """Check if there's an active presentation open."""
        try:
            if self._ppt_app:
                # Check if there are any presentations open
                if self._ppt_app.Presentations.Count > 0:
                    # Check if there's an active window
                    if self._ppt_app.ActiveWindow:
                        return True
            return False
        except Exception as e:
            log.debug(f"No active presentation: {e}")
            return False

    def _verify_connection(self):
        """Verify COM connection is alive."""
        try:
            # Simple test - access ActivePresentation
            _ = self._ppt_app.ActivePresentation.Name
            return True
        except Exception:
            # Reconnect
            return self._connect_to_powerpoint()

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
                ui.message("Switched to Normal view")
                return True
            else:
                log.debug("Already in Normal view")
        except Exception as e:
            log.debug(f"Failed to switch view: {e}")
        return False
