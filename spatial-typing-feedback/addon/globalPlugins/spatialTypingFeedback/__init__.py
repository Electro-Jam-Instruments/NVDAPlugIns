# globalPlugins/spatialTypingFeedback
# Separates typing feedback from NVDA's main speech by stereo position.
#
# Version: 0.1.0
#
# Three streams (docs/user-requirements.md):
#   Main   - NVDA's main voice and completed words   centre,     100%
#   Chars  - typed characters                        +35 right,   50%
#   Errors - the typing-error alert                  -35 left,    50%
#
# Values are hardcoded in _streams.py on purpose until they have been judged by ear.
# See docs/architecture-decisions.md Decision D11.

import os

import addonHandler
import config
import globalPluginHandler
import globalVars
import synthDriverHandler
import ui
from logHandler import log
from scriptHandler import script

from ._audio import PannedPlayer
from ._echo import EchoInterceptor
from ._errors import ErrorAlertInterceptor
from ._streams import CHAR_ECHO_STREAM, ERROR_ALERT_STREAM, WORD_ECHO_STREAM, panAndVolume
from ._voice import OneCoreVoice

try:
    addonHandler.initTranslation()
except Exception:  # noqa: BLE001
    pass

try:
    _
except NameError:
    # Translations not installed (e.g. loaded from the developer scratchpad).
    def _(text):
        return text

VERSION = "0.1.0"

log.info("Spatial Typing Feedback: loading plugin v%s" % VERSION)


def _soundSplitActive():
    """True if NVDA's Sound Split is on.

    Sound Split sets channel volume on the whole NVDA process audio session, one level
    above our per-player volumes, so the two multiply. We do not try to reason about
    individual states - the features are mutually exclusive (Decision D7).
    """
    try:
        return config.conf["audio"]["soundSplitState"] != 0
    except Exception:  # noqa: BLE001
        return False


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    """Routes typing feedback to distinct stereo positions using a secondary voice."""

    def __init__(self):
        super().__init__()
        self._active = False
        self._dormantReason = None
        self._charVoice = None
        self._wordVoice = None
        self._errorPlayer = None
        self._echo = None
        self._errors = None
        try:
            self._start()
        except Exception:  # noqa: BLE001
            log.error("Spatial Typing Feedback: failed to start", exc_info=True)
            self._stop()

    # -- lifecycle --------------------------------------------------------------

    def _start(self):
        if _soundSplitActive():
            self._dormantReason = _(
                # Translators: Reported when the add-on cannot run because Sound Split is on.
                "Spatial typing feedback is off because NVDA's sound split is enabled",
            )
            log.warning("Spatial Typing Feedback: dormant, sound split is enabled")
            return

        # Errors need no voice at all - just a wave file at a position.
        pan, vol = panAndVolume(ERROR_ALERT_STREAM)
        self._errorPlayer = PannedPlayer()
        self._errorPlayer.setPosition(pan, vol)
        self._errors = ErrorAlertInterceptor(self._errorPlayer)
        self._errors.install()

        # Characters and completed words each get their own OneCore instance, because
        # one space fires both at once and they need different positions and levels
        # at the same instant.
        self._charVoice = self._makeVoice("chars", CHAR_ECHO_STREAM)
        self._wordVoice = self._makeVoice("words", WORD_ECHO_STREAM)

        if self._charVoice or self._wordVoice:
            self._echo = EchoInterceptor(onWord=self._speakWord, onChar=self._speakChar)
            if not self._echo.install():
                self._echo = None
        else:
            log.warning(
                "Spatial Typing Feedback: no secondary voice available; typing echo left "
                "with NVDA's main voice",
            )

        self._active = True
        log.info("Spatial Typing Feedback: active")

    def _makeVoice(self, name, stream):
        voice = OneCoreVoice(name)
        if not voice.initialize():
            return None
        self._matchMainVoice(voice)
        pan, vol = panAndVolume(stream)
        voice.setPosition(pan, vol)
        return voice

    def _matchMainVoice(self, voice):
        """Mirror the main synth's voice and settings (R6).

        Separation is positional, not timbral - the streams should sound the same and
        be told apart by where they are. Rate boost matters most: an echo capped at
        normal speed would still be arriving after a boosted main voice had moved on.
        """
        try:
            synth = synthDriverHandler.getSynth()
        except Exception:  # noqa: BLE001
            return
        if synth is None:
            return
        try:
            if getattr(synth, "name", None) == "oneCore":
                voiceId = getattr(synth, "voice", None)
                if voiceId:
                    voice.setVoiceById(voiceId)
            rateBoost = bool(getattr(synth, "rateBoost", False))
            voice.setRate(int(getattr(synth, "rate", 50)), rateBoost)
            if synth.isSupported("pitch"):
                voice.setPitch(int(getattr(synth, "pitch", 50)))
            if synth.isSupported("volume"):
                voice.setVolume(int(getattr(synth, "volume", 100)))
        except Exception:  # noqa: BLE001
            log.debugWarning("Spatial Typing Feedback: could not mirror main voice", exc_info=True)

    def _stop(self):
        if self._echo is not None:
            self._echo.uninstall()
            self._echo = None
        if self._errors is not None:
            self._errors.uninstall()
            self._errors = None
        for voice in (self._charVoice, self._wordVoice):
            if voice is not None:
                voice.terminate()
        self._charVoice = None
        self._wordVoice = None
        if self._errorPlayer is not None:
            self._errorPlayer.terminate()
            self._errorPlayer = None
        self._active = False

    def terminate(self):
        try:
            self._stop()
        except Exception:  # noqa: BLE001
            log.error("Spatial Typing Feedback: error during termination", exc_info=True)
        log.info("Spatial Typing Feedback: plugin terminated")
        super().terminate()

    # -- speech routing ---------------------------------------------------------

    def _speakWord(self, text):
        if self._wordVoice is not None:
            self._wordVoice.speak(text)

    def _speakChar(self, text):
        if self._charVoice is not None:
            self._charVoice.speak(text)

    # -- scripts ----------------------------------------------------------------

    @script(
        # Translators: Input help mode message for a command that toggles the add-on.
        description=_("Toggles spatial typing feedback on and off"),
        gesture="kb:NVDA+shift+s",
        category="Spatial Typing Feedback",
    )
    def script_toggle(self, gesture):
        if self._active:
            self._stop()
            # Translators: Reported when the add-on is turned off.
            ui.message(_("Spatial typing feedback off"))
            return
        self._dormantReason = None
        try:
            self._start()
        except Exception:  # noqa: BLE001
            log.error("Spatial Typing Feedback: failed to start", exc_info=True)
            self._stop()
        if self._dormantReason:
            ui.message(self._dormantReason)
        elif self._active:
            # Translators: Reported when the add-on is turned on.
            ui.message(_("Spatial typing feedback on"))

    @script(
        # Translators: Input help mode message for a command that plays a test sound
        # at each stream position.
        description=_("Plays a test phrase at each stream position"),
        gesture="kb:NVDA+shift+p",
        category="Spatial Typing Feedback",
    )
    def script_testPositions(self, gesture):
        """Hear all three streams in turn, so the values can be judged by ear."""
        if self._dormantReason:
            ui.message(self._dormantReason)
            return
        parts = []
        if self._wordVoice is not None:
            # Translators: Spoken by the test command, from the centre stream.
            self._wordVoice.speak(_("main, centre"))
            parts.append("main")
        if self._charVoice is not None:
            # Translators: Spoken by the test command, from the character stream.
            self._charVoice.speak(_("characters, right"))
            parts.append("chars")
        if self._errorPlayer is not None:
            errorWav = os.path.join(globalVars.appDir, "waves", "textError.wav")
            if os.path.isfile(errorWav):
                self._errorPlayer.feedWaveFile(errorWav)
                parts.append("errors")
        if not parts:
            # Translators: Reported when the test command has nothing available to play.
            ui.message(_("No streams available"))
