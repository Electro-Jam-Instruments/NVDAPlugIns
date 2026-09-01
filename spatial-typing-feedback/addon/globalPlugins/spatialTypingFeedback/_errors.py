# Positioning NVDA's typing-error alert.
#
# See docs/architecture.md.
#
# NVDA reports spelling errors three ways. Only ONE of them is a typing event:
#
#   1. While typing  -> playWaveFile(textError.wav), isSpeechWaveFileCommand=False  <- ours
#   2. While reading -> the same wav, but as a WaveFileCommand inside a speech
#                       sequence, so isSpeechWaveFileCommand=True
#   3. While reading -> NVDA speaks "spelling error" as part of the sequence
#
# 2 and 3 annotate text the main voice is reading. Moving them would divorce the
# annotation from the words it describes, so they are left exactly where they are.

import os

import nvwave
from logHandler import log

#: The wave NVDA plays when you finish a misspelled word.
TEXT_ERROR_WAV = "texterror.wav"


class ErrorAlertInterceptor(object):
    """Plays the typing-error alert at our own position instead of NVDA's."""

    def __init__(self, player, onWavePath=None):
        self._player = player
        self._onWavePath = onWavePath
        self._installed = False

    def install(self):
        if self._installed:
            return True
        decider = getattr(nvwave, "decide_playWaveFile", None)
        if decider is None:
            log.error(
                "Spatial Typing Feedback: nvwave.decide_playWaveFile not available in this "
                "NVDA version; the error alert will not be positioned",
            )
            return False
        try:
            decider.register(self._decide)
        except Exception:  # noqa: BLE001
            log.error("Spatial Typing Feedback: could not register wave file hook", exc_info=True)
            return False
        self._installed = True
        log.info("Spatial Typing Feedback: error alert interception installed")
        return True

    def uninstall(self):
        if not self._installed:
            return
        try:
            nvwave.decide_playWaveFile.unregister(self._decide)
        except Exception:  # noqa: BLE001
            log.debugWarning("Error unregistering wave file hook", exc_info=True)
        self._installed = False

    def _decide(self, fileName=None, asynchronous=True, isSpeechWaveFileCommand=False, **kwargs):
        """Return False to cancel NVDA's playback because we are handling it."""
        try:
            if isSpeechWaveFileCommand:
                # Part of a speech sequence - belongs with the main voice.
                return True
            if not fileName or os.path.basename(fileName).lower() != TEXT_ERROR_WAV:
                return True
            if self._onWavePath is not None:
                # Remember the path NVDA used. It is correct by construction,
                # which one we build ourselves is not guaranteed to be.
                self._onWavePath(fileName)
            self._player.feedWaveFile(fileName)
            return False
        except Exception:  # noqa: BLE001
            log.error("Spatial Typing Feedback: error handling wave file", exc_info=True)
            # Let NVDA play it normally rather than swallowing the alert.
            return True
