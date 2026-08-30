# Intercepting NVDA's typing echo.
#
# See docs/architecture-decisions.md Decision D4. Character echo (R1) and word echo (R3)
# are two branches of ONE function, speech.speakTypedCharacters, so this is one hook.

import speech
import speech.speech as speechModule
from logHandler import log


class EchoInterceptor(object):
    """Redirects NVDA's typing echo to our panned voices.

    Rather than reimplementing NVDA's decision logic, we let NVDA's own
    speakTypedCharacters run and swap out the two functions it uses to produce output.

    This matters. speakTypedCharacters owns the typed-word buffer, honours protected
    fields, respects the suppression counter and handles the TypingEcho modes - and
    terminals call speech.clearTypedWordBuffer() from elsewhere. Reimplementing all of
    that would mean duplicating state we cannot see and racing with code we do not
    control. Capturing the two outputs leaves every decision with NVDA and takes only
    what it decided to say.

    The swap is in place only for the duration of one synchronous call on the main
    thread, so the window in which another caller could be affected is vanishingly
    small.
    """

    def __init__(self, onWord, onChar):
        self._onWord = onWord
        self._onChar = onChar
        self._original = None
        self._installed = False

    def install(self):
        if self._installed:
            return True
        original = getattr(speech, "speakTypedCharacters", None)
        if original is None or not callable(original):
            log.error(
                "Spatial Typing Feedback: speech.speakTypedCharacters not found; "
                "typing echo will be left alone",
            )
            return False
        if not hasattr(speechModule, "speakText") or not hasattr(speechModule, "speakSpelling"):
            log.error(
                "Spatial Typing Feedback: speech.speech.speakText/speakSpelling not found; "
                "typing echo will be left alone",
            )
            return False
        self._original = original
        speech.speakTypedCharacters = self._wrapped
        self._installed = True
        log.info("Spatial Typing Feedback: typing echo interception installed")
        return True

    def uninstall(self):
        if not self._installed:
            return
        # Only restore if nobody else has patched over us in the meantime.
        if getattr(speech, "speakTypedCharacters", None) is self._wrapped:
            speech.speakTypedCharacters = self._original
        else:
            log.warning(
                "Spatial Typing Feedback: speakTypedCharacters was replaced by something "
                "else; leaving it alone",
            )
        self._installed = False
        self._original = None

    def _wrapped(self, ch):
        """Run NVDA's logic, capture its output, speak it ourselves."""
        origSpeakText = speechModule.speakText
        origSpeakSpelling = speechModule.speakSpelling
        speechModule.speakText = self._captureWord
        speechModule.speakSpelling = self._captureChar
        try:
            self._original(ch)
        except Exception:  # noqa: BLE001
            log.error("Spatial Typing Feedback: error in typing echo", exc_info=True)
        finally:
            speechModule.speakText = origSpeakText
            speechModule.speakSpelling = origSpeakSpelling

    def _captureWord(self, text, *args, **kwargs):
        try:
            self._onWord(text)
        except Exception:  # noqa: BLE001
            log.error("Spatial Typing Feedback: error speaking word echo", exc_info=True)

    def _captureChar(self, text, *args, **kwargs):
        try:
            self._onChar(text)
        except Exception:  # noqa: BLE001
            log.error("Spatial Typing Feedback: error speaking character echo", exc_info=True)
