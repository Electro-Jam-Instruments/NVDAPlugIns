# globalPlugins/windowsDictationSilence.py
# Auto-silence NVDA during Windows Voice Typing (Win+H)
#
# v0.0.1: Initial implementation - focus-based detection (didn't work)
# v0.0.2: Win+H hotkey interception + timer-based window polling
# v0.0.3: Keypress interception approach (no timers)
# v0.0.4: Fix gesture passthrough order - send Win+H before installing filter
# v0.0.5: Fix repeated Win+H - don't skip Win+H in gesture filter
# v0.0.6: Fix API - use speech.getState().speechMode instead of getSpeechMode()
#
# See docs/ folder for full documentation.

import globalPluginHandler
import speech
import inputCore
import logging
from scriptHandler import script

log = logging.getLogger(__name__)
log.info("Windows Dictation Silence: Loading plugin v0.0.6")


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    """Auto-silence NVDA during Windows Voice Typing.

    Approach:
    1. Intercept Win+H to detect Voice Typing starting
    2. Set speech mode to OFF
    3. Install gesture filter to catch ANY subsequent keypress
    4. On any keypress: restore speech, remove filter, pass key through

    This works because Voice Typing closes on any keyboard input.
    """

    def __init__(self):
        super().__init__()
        self._previous_speech_mode = None
        self._voice_typing_active = False
        self._gesture_filter_installed = False
        log.info("Windows Dictation Silence: Plugin initialized")

    def terminate(self):
        """Clean up on plugin unload."""
        # Remove gesture filter if active
        if self._gesture_filter_installed:
            self._remove_gesture_filter()

        # Restore speech if we're being unloaded while voice typing is active
        if self._voice_typing_active and self._previous_speech_mode is not None:
            speech.setSpeechMode(self._previous_speech_mode)

        log.info("Windows Dictation Silence: Plugin terminated")

    @script(
        description="Toggle Windows Voice Typing with auto-silence",
        gesture="kb:windows+h",
    )
    def script_toggleVoiceTyping(self, gesture):
        """Intercept Win+H to manage speech around Voice Typing."""
        log.info(f"Windows Dictation Silence: Win+H pressed (active={self._voice_typing_active})")

        # Always pass through Win+H to Windows
        gesture.send()

        # If not active, start voice typing mode (filter will handle the close)
        if not self._voice_typing_active:
            self._start_voice_typing_mode()
        # If active, filter already called _end_voice_typing_mode() before we got here

    def _start_voice_typing_mode(self):
        """Enter voice typing mode - silence speech and watch for close."""
        self._previous_speech_mode = speech.getState().speechMode
        speech.setSpeechMode(speech.SpeechMode.off)
        self._voice_typing_active = True
        self._install_gesture_filter()
        log.info(f"Windows Dictation Silence: Speech OFF (was {self._previous_speech_mode})")

    def _end_voice_typing_mode(self):
        """Exit voice typing mode - restore speech."""
        self._remove_gesture_filter()
        self._voice_typing_active = False

        if self._previous_speech_mode is not None:
            speech.setSpeechMode(self._previous_speech_mode)
            log.info(f"Windows Dictation Silence: Speech restored to {self._previous_speech_mode}")
            self._previous_speech_mode = None

    def _install_gesture_filter(self):
        """Install filter to intercept any keypress."""
        if self._gesture_filter_installed:
            return

        try:
            inputCore.decide_executeGesture.register(self._gesture_filter)
            self._gesture_filter_installed = True
            log.debug("Windows Dictation Silence: Gesture filter installed")
        except Exception as e:
            log.error(f"Windows Dictation Silence: Failed to install gesture filter - {e}")

    def _remove_gesture_filter(self):
        """Remove the gesture filter."""
        if not self._gesture_filter_installed:
            return

        try:
            inputCore.decide_executeGesture.unregister(self._gesture_filter)
            self._gesture_filter_installed = False
            log.debug("Windows Dictation Silence: Gesture filter removed")
        except Exception as e:
            log.error(f"Windows Dictation Silence: Failed to remove gesture filter - {e}")

    def _gesture_filter(self, gesture, *args, **kwargs):
        """Filter called for every gesture while Voice Typing is active.

        Any keypress means Voice Typing is closing, so restore speech.
        This includes Win+H (which toggles Voice Typing off).

        Returns:
            True to allow the gesture to proceed
        """
        if not self._voice_typing_active:
            return True

        # Get gesture identifier for logging
        gesture_id = getattr(gesture, 'identifiers', [''])[0] if hasattr(gesture, 'identifiers') else ''

        # ANY key (including Win+H) means Voice Typing is closing
        log.info(f"Windows Dictation Silence: Key pressed ({gesture_id}) - restoring speech")
        self._end_voice_typing_mode()

        return True  # Allow the gesture to proceed


log.info("Windows Dictation Silence: Plugin module loaded")
