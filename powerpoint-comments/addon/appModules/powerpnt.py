# appModules/powerpnt.py
# PowerPoint Comments NVDA Addon - Phase 1: Foundation + View Management
#
# This module extends NVDA's built-in PowerPoint support to add
# comment navigation features.

# First, inherit all built-in PowerPoint support
from nvdaBuiltin.appModules.powerpnt import *

import appModuleHandler
from comtypes.client import GetActiveObject
import ui
import logging

log = logging.getLogger(__name__)


class AppModule(appModuleHandler.AppModule):
    """Enhanced PowerPoint with comment navigation."""

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
        log.info("PowerPoint Comments addon initialized (v0.0.1)")

    def event_appModule_gainFocus(self):
        """Called when PowerPoint gains focus."""
        log.debug("event_appModule_gainFocus fired")
        self._connect_to_powerpoint()
        self._ensure_normal_view()

    def _connect_to_powerpoint(self):
        """Connect to running PowerPoint instance."""
        try:
            self._ppt_app = GetActiveObject("PowerPoint.Application")
            log.debug("Connected to PowerPoint COM")
            return True
        except Exception as e:
            log.error(f"Failed to connect to PowerPoint: {e}")
            self._ppt_app = None
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
            log.error(f"Failed to get view type: {e}")
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
            log.error(f"Failed to switch view: {e}")
        return False
