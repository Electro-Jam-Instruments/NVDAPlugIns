# globalPlugins/spatialTypingFeedback
# Separates typing feedback from NVDA's main speech by stereo position.
#
# Three streams, described in full in docs/architecture.md:
#
#   Main    NVDA's own voice, and completed words     centre
#   Chars   each typed character                      right, quieter, our own voice
#   Errors  the typing-error alert                    left, quieter, a wave file
#
# Only the character echo needs a voice of its own. Word echo goes back to NVDA, which
# already speaks it at the user's rate and rate boost.

import ctypes
import os
import sys
import time

import addonHandler
import api
import config
import globalPluginHandler
import globalVars
import inputCore
import scriptHandler
import speech
import synthDriverHandler
import ui
from logHandler import log
from scriptHandler import script

from . import _config, _ring, _streams
from ._audio import PannedPlayer
from ._echo import EchoInterceptor
from ._errors import ErrorAlertInterceptor
from ._streams import (
    CHAR_ECHO_STREAM,
    ERROR_ALERT_STREAM,
    ONECORE_MAX_SPEAKING_RATE,
    ONECORE_MIN_SPEAKING_RATE,
    WORD_ECHO_USES_MAIN_VOICE,
    Stream,
)
from ._voice import SapiVoice
from ._winrt import WinRTVoice

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

_config.initialize()


def _soundSplitActive():
    """True if NVDA's Sound Split is on.

    Sound Split sets channel volume on the whole NVDA process audio session, one level
    above our per-player volumes, so the two multiply. We do not try to reason about
    individual states - the features are mutually exclusive (see docs/architecture.md).
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
        self._voiceError = None
        self.engineName = None
        self._annotationVoice = None
        self._annotationMode = 0
        self._cancelHookInstalled = False
        self._speechFilterInstalled = False
        self._gestureHookInstalled = False
        self._suppressUntil = None
        self._suppressText = None
        #: NVDA's wording for the spelling and grammar markers, read once.
        self._annotationTexts = None
        self._originalRing = None
        self.selectedStreamIndex = 0
        self._lastErrorWave = None
        self._voice = None
        self._charPlayer = None
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
        self._voiceError = None
        if _soundSplitActive():
            self._dormantReason = _(
                # Translators: Reported when the add-on cannot run because Sound Split is on.
                "Spatial typing feedback is off because NVDA's sound split is enabled",
            )
            log.warning("Spatial Typing Feedback: dormant, sound split is enabled")
            return

        # Errors need no voice at all - just a wave file at a position.
        self._errorPlayer = self._makePlayer("errors", ERROR_ALERT_STREAM)
        self._errors = ErrorAlertInterceptor(
            self._errorPlayer,
            onWavePath=self._noteErrorWave,
            positionSequenceSounds=bool(self._annotationMode & 2),
            shouldPlay=self.shouldSpeak,
        )
        self._errors.install()

        # One OneCore token, one player per stream. Each utterance carries its
        # destination, so the two echo streams stay at their own positions without
        # depending on a second token being distinct from the first.
        self._charPlayer = self._makePlayer("chars", CHAR_ECHO_STREAM)

        mode = _config.get("annotationMode", 0)
        self._annotationMode = mode
        if mode:
            log.info("Spatial Typing Feedback: annotation mode %d" % mode)
        if mode & 1:
            self._annotationVoice = self._makeAnnotationVoice()
            if self._annotationVoice is not None:
                self.setUpVoicedStream(Stream.ANNOTATIONS)
        voice = self._makeSecondaryVoice()
        if voice is not None:
            self._voice = voice
            self.setUpVoicedStream(Stream.CHARS)
            # Completed words stay with NVDA's main voice: it already has the user's
            # rate and rate boost, and the word echo was never the stream in the way.
            onWord = None if WORD_ECHO_USES_MAIN_VOICE else self._speakWord
            self._echo = EchoInterceptor(onChar=self._speakChar, onWord=onWord)
            if not self._echo.install():
                self._echo = None
            self._installCancelHook()
            self._installSpeechFilter()
            self._installGestureHook()
            self._originalRing = _ring.install(self)
        else:
            log.warning(
                "Spatial Typing Feedback: no secondary voice available; typing echo left "
                "with NVDA's main voice",
            )

        self._active = True
        log.info("Spatial Typing Feedback: active")

    def _makeSecondaryVoice(self):
        """Real OneCore if we can get an independent instance, SAPI 5 otherwise.

        OneCore is strongly preferred: it is the same engine as the main voice, with the
        same voices and the same rate boost, so the two streams match and the echo keeps
        up. SAPI 5 works everywhere but is an older voice build and tops out around
        three times normal speed.
        """
        voice = WinRTVoice("echo")
        if voice.initialize():
            self.engineName = "Windows voices"
            return voice
        log.warning(
            "Spatial Typing Feedback: Windows SpeechSynthesizer unavailable (%s); "
            "falling back to SAPI 5" % voice.lastError,
        )
        voice.terminate()
        voice = SapiVoice("echo")
        if voice.initialize():
            self.engineName = "SAPI 5"
            return voice
        self._voiceError = _(
            # Translators: Reported when the add-on could not create its second voice.
            "Secondary voice unavailable, so typing echo is unchanged. "
            "See the NVDA log for details",
        )
        self.engineName = None
        return None

    def _makeAnnotationVoice(self):
        """A separate synthesizer for the spelling and grammar annotations.

        Its own instance rather than sharing the character voice, so an annotation can
        play at the same time as typing rather than queueing behind it. Possible only
        because we activate SpeechSynthesizer ourselves; NVDA's wrapper allows one.

        Failure is not fatal - without it the annotations simply stay with the main
        voice, which is what NVDA does today.
        """
        voice = WinRTVoice("errors")
        if not voice.initialize():
            log.warning(
                "Spatial Typing Feedback: no voice for error annotations (%s); they stay "
                "with the main voice" % voice.lastError,
            )
            return None
        pan, _vol = _config.panAndVolume(Stream.ANNOTATIONS)
        voice.player.setPosition(pan, self.streamVolume(Stream.ANNOTATIONS))
        # Match the main voice: these are NVDA's own words, not ours. Sharing a voice
        # with NVDA is fine - three instances on one voice all produce audio.
        synth = self._synth()
        if synth is not None:
            voiceId = getattr(synth, "voice", None)
            if voiceId:
                voice.matchVoice(voiceId)
        return voice

    def _makePlayer(self, name, stream):
        pan, vol = _config.panAndVolume(stream)
        player = PannedPlayer(name)
        player.setPosition(pan, vol)
        return player

    # -- live tuning -------------------------------------------------------------

    @property
    def selectedStream(self):
        return _ring.STREAMS[self.selectedStreamIndex]

    def _playerFor(self, stream):
        """The player a stream's audio comes out of.

        Voiced streams own their player through the synthesizer, so ask the voice rather
        than keeping a second mapping that can fall out of step - leaving one stream out
        of this table silently stopped its volume and pan from being applied at all.
        """
        if stream is Stream.CHARS:
            return self._charPlayer
        if stream is Stream.ERRORS:
            return self._errorPlayer
        voice = self.voiceFor(stream)
        return voice.player if voice is not None else None

    def nvdaSoundVolume(self):
        """The level NVDA plays its own sounds at.

        Mirrors nvwave's own rule, including soundVolumeFollowsVoice, so the relocated
        error alert stays at the level the user already chose for NVDA's sounds instead
        of needing to be set twice.
        """
        try:
            if config.conf["audio"]["soundVolumeFollowsVoice"]:
                synth = self._synth()
                if synth is not None and synth.isSupported("volume"):
                    return int(synth.volume)
            return int(config.conf["audio"]["soundVolume"])
        except Exception:  # noqa: BLE001
            return 100

    def streamVolume(self, stream):
        """Configured level, resolving the defaults that are relative to the main voice.

        The alert follows NVDA's own sound volume. The voiced side streams start a fixed
        step below the main voice, then become independent once set.
        """
        _pan, vol = _config.panAndVolume(stream)
        if vol != _config.UNSET:
            return vol
        if stream is Stream.ERRORS:
            return self.nvdaSoundVolume()
        return max(0, min(100, self.getMainVolume() + _streams.SIDE_VOLUME_OFFSET))

    def isFollowingNvdaVolume(self, stream):
        if stream is not Stream.ERRORS:
            return False
        _pan, vol = _config.panAndVolume(stream)
        return vol == _config.UNSET

    def _applyStream(self, stream):
        player = self._playerFor(stream)
        if player is None:
            return
        pan, _vol = _config.panAndVolume(stream)
        volume = self.streamVolume(stream)
        player.setPosition(pan, volume)

    def setStreamPan(self, stream, pan):
        panKey, _volKey = _config.STREAM_KEYS[stream]
        _config.set(panKey, pan)
        self._applyStream(stream)
        self._preview(stream)

    def setStreamVolume(self, stream, volume):
        """Set a stream's level. For errors, UNSET means follow NVDA's sound volume."""
        _panKey, volKey = _config.STREAM_KEYS[stream]
        volume = int(volume)
        if volume < 0:
            volume = _config.UNSET if stream is Stream.ERRORS else 0
        _config.set(volKey, volume)
        self._applyStream(stream)
        self._preview(stream)

    def _noteErrorWave(self, path):
        self._lastErrorWave = path

    def _errorWavePath(self):
        """Where NVDA's typing-error wave lives.

        Remembered from the last real alert when we have seen one - that path came from
        NVDA itself and is correct by construction. Otherwise build it, trying more than
        one root, because globalVars.appDir is not the only way NVDA describes its own
        directory and getting it wrong fails silently.
        """
        if self._lastErrorWave and os.path.isfile(self._lastErrorWave):
            return self._lastErrorWave
        roots = []
        for getter in (
            lambda: globalVars.appDir,
            lambda: os.path.dirname(sys.executable),
        ):
            try:
                roots.append(getter())
            except Exception:  # noqa: BLE001
                pass
        for root in roots:
            if not root:
                continue
            candidate = os.path.join(root, "waves", "textError.wav")
            if os.path.isfile(candidate):
                self._lastErrorWave = candidate
                return candidate
        log.warning(
            "Spatial Typing Feedback: could not find textError.wav; looked under %r" % roots,
        )
        return None

    def _preview(self, stream):
        """Play the error alert so a change to that stream can be heard.

        ONLY for the error stream. The character stream needs nothing here: its ring
        announcements are already spoken by the character voice itself, at its own
        position, speed and pitch - so the announcement is the sample. Adding a separate
        preview word meant hearing the old settings and then the new ones, which is worse
        than useless when you are trying to judge a change.
        """
        if stream is not Stream.ERRORS:
            return
        try:
            path = self._errorWavePath()
            if self._errorPlayer is not None and path:
                self._errorPlayer.feedWaveFile(path)
        except Exception:  # noqa: BLE001
            log.error("Spatial Typing Feedback: error playing preview", exc_info=True)

    # -- proxies to NVDA's own voice, so one selector covers every stream ----------

    def _synth(self):
        try:
            return synthDriverHandler.getSynth()
        except Exception:  # noqa: BLE001
            return None

    def getMainVolume(self):
        synth = self._synth()
        try:
            return int(getattr(synth, "volume", 100)) if synth else 100
        except Exception:  # noqa: BLE001
            return 100

    def setMainVolume(self, value):
        synth = self._synth()
        if synth is None:
            return
        try:
            synth.volume = value
            synth.saveSettings()
        except Exception:  # noqa: BLE001
            log.debugWarning("Could not set main volume", exc_info=True)

    def getMainRate(self):
        synth = self._synth()
        try:
            return int(getattr(synth, "rate", 50)) if synth else 50
        except Exception:  # noqa: BLE001
            return 50

    def setMainRate(self, value):
        synth = self._synth()
        if synth is None:
            return
        try:
            synth.rate = value
            synth.saveSettings()
        except Exception:  # noqa: BLE001
            log.debugWarning("Could not set main rate", exc_info=True)

    def getMainVoices(self):
        """[(id, displayName)] for the main synth."""
        synth = self._synth()
        if synth is None:
            return []
        try:
            return [(v.id, v.displayName) for v in synth.availableVoices.values()]
        except Exception:  # noqa: BLE001
            return []

    def getMainVoiceIndex(self):
        synth = self._synth()
        if synth is None:
            return 0
        try:
            current = synth.voice
            for i, entry in enumerate(self.getMainVoices()):
                if entry[0] == current:
                    return i
        except Exception:  # noqa: BLE001
            pass
        return 0

    def setMainVoiceIndex(self, index):
        synth = self._synth()
        voices = self.getMainVoices()
        if synth is None or not voices:
            return
        try:
            synthDriverHandler.changeVoice(synth, voices[index][0])
            synth.saveSettings()
        except Exception:  # noqa: BLE001
            log.debugWarning("Could not set main voice", exc_info=True)

    # -- per-stream voice settings ------------------------------------------------
    #
    # Generic rather than one set of accessors per stream. Two voiced streams already
    # share every setting; duplicating the accessors would guarantee they drift apart,
    # and a fifth stream would mean another ten methods.

    def voiceFor(self, stream):
        """The synthesizer behind a stream, or None if it has no voice."""
        return {
            Stream.CHARS: self._voice,
            Stream.ANNOTATIONS: self._annotationVoice,
        }.get(stream)

    def streamHasVoice(self, stream):
        return self.voiceFor(stream) is not None

    def _voiceSetting(self, stream, name, default):
        value = _config.get(_config.voiceKey(stream, name), default)
        return default if value == _config.UNSET else value

    def getStreamRate(self, stream):
        return self._voiceSetting(stream, "Rate", 50)

    def setStreamRate(self, stream, value):
        _config.set(_config.voiceKey(stream, "Rate"), max(0, min(100, int(value))))
        self.applyVoiceSettings(stream)

    def getStreamPitch(self, stream):
        return self._voiceSetting(stream, "Pitch", 50)

    def setStreamPitch(self, stream, value):
        _config.set(_config.voiceKey(stream, "Pitch"), max(0, min(100, int(value))))
        self.applyVoiceSettings(stream)

    def getStreamRateBoost(self, stream):
        return bool(_config.get(_config.voiceKey(stream, "RateBoost"), True))

    def setStreamRateBoost(self, stream, enable):
        _config.set(_config.voiceKey(stream, "RateBoost"), bool(enable))
        # The percentage is kept and its meaning widens or narrows, as NVDA does.
        self.applyVoiceSettings(stream)

    def streamSupportsPunctuation(self, stream):
        voice = self.voiceFor(stream)
        return bool(getattr(voice, "supportsPunctuationSilence", False))

    def getStreamPunctuation(self, stream):
        return bool(_config.get(_config.voiceKey(stream, "Punctuation"), True))

    def setStreamPunctuation(self, stream, enable):
        _config.set(_config.voiceKey(stream, "Punctuation"), bool(enable))
        voice = self.voiceFor(stream)
        if voice is not None and hasattr(voice, "setPunctuationSilence"):
            voice.setPunctuationSilence(enable)

    def getStreamVoices(self, stream):
        """[(id, displayName)] for a stream's synthesizer."""
        voice = self.voiceFor(stream)
        if voice is None or not hasattr(voice, "getAvailableVoiceIds"):
            return []
        try:
            return [
                (vid, self._voiceDisplayName(voice, vid))
                for vid, _index in voice.getAvailableVoiceIds()
            ]
        except Exception:  # noqa: BLE001
            return []

    def _voiceDisplayName(self, voice, voiceId):
        """A speakable name for a voice ID.

        The engine reports proper display names; the raw IDs are registry paths and
        unusable read aloud.
        """
        if hasattr(voice, "getVoiceDisplayName"):
            name = voice.getVoiceDisplayName(voiceId)
            if name:
                return name
        for entry in self.getMainVoices():
            if entry[0] == voiceId:
                return entry[1]
        return voiceId.rstrip(chr(92)).split(chr(92))[-1]

    def getStreamVoiceIndex(self, stream):
        wanted = _config.get(_config.voiceKey(stream, "Voice"), "")
        for i, entry in enumerate(self.getStreamVoices(stream)):
            if entry[0] == wanted:
                return i
        return 0

    def setStreamVoiceIndex(self, stream, index):
        voices = self.getStreamVoices(stream)
        voice = self.voiceFor(stream)
        if not voices or voice is None:
            return
        voiceId = voices[index][0]
        try:
            voice.matchVoice(voiceId)
            _config.set(_config.voiceKey(stream, "Voice"), voiceId)
            # Changing voice can reset prosody, so put the settings back.
            self.applyVoiceSettings(stream)
        except Exception:  # noqa: BLE001
            log.debugWarning("Could not set the voice for %s" % stream, exc_info=True)

    def applyVoiceSettings(self, stream):
        """Push a stream's stored speed and pitch onto its synthesizer."""
        voice = self.voiceFor(stream)
        if voice is None:
            return
        try:
            voice.setRatePercent(self.getStreamRate(stream), self.getStreamRateBoost(stream))
            voice.setPitchPercent(self.getStreamPitch(stream))
        except Exception:  # noqa: BLE001
            log.error("Spatial Typing Feedback: could not apply voice settings", exc_info=True)

    def _seedStreamFromMain(self, stream, synth):
        """Give a voiced stream a sensible starting point, once.

        Only when nothing has been chosen. After that it is independent - the point is
        that it can be set where you want without the main voice dragging it back.
        """
        rateKey = _config.voiceKey(stream, "Rate")
        if _config.get(rateKey, _config.UNSET) == _config.UNSET:
            try:
                section = config.conf["speech"].get(getattr(synth, "name", "") or "")
                rate = int(section.get("rate", 50)) if section else 50
                rateBoost = bool(section.get("rateBoost", False)) if section else False
            except Exception:  # noqa: BLE001
                rate, rateBoost = 50, False
            # Our scale always spans the full range, so a non-boosted main rate has to be
            # converted rather than copied straight across.
            maxRate = 6.0 if rateBoost else 1.5
            raw = float(rate) / 100 * (maxRate - 0.5) + 0.5
            seeded = int(round((raw - 0.5) / (6.0 - 0.5) * 100))
            _config.set(rateKey, max(0, min(100, seeded)))
        volKey = _config.STREAM_KEYS[stream][1]
        if _config.get(volKey, _config.UNSET) == _config.UNSET:
            _config.set(
                volKey,
                max(0, min(100, self.getMainVolume() + _streams.SIDE_VOLUME_OFFSET)),
            )
        pitchKey = _config.voiceKey(stream, "Pitch")
        if _config.get(pitchKey, _config.UNSET) == _config.UNSET:
            try:
                _config.set(pitchKey, int(getattr(synth, "pitch", 50)))
            except Exception:  # noqa: BLE001
                _config.set(pitchKey, 50)

    def setUpVoicedStream(self, stream):
        """Prepare a voiced stream: its voice, speed, pitch and punctuation."""
        voice = self.voiceFor(stream)
        if voice is None:
            return
        synth = self._synth()
        if synth is not None:
            self._seedStreamFromMain(stream, synth)
        saved = _config.get(_config.voiceKey(stream, "Voice"), "")
        if saved:
            voice.matchVoice(saved)
        elif synth is not None:
            voiceId = getattr(synth, "voice", None)
            if voiceId:
                voice.matchVoice(voiceId)
                _config.set(_config.voiceKey(stream, "Voice"), voiceId)
        self.applyVoiceSettings(stream)
        if hasattr(voice, "setPunctuationSilence"):
            voice.setPunctuationSilence(self.getStreamPunctuation(stream))

    def playerForVoicedStream(self, stream):
        """Where a voiced stream's speech goes."""
        if stream is Stream.CHARS:
            return self._charPlayer
        voice = self.voiceFor(stream)
        return voice.player if voice is not None else None

    def announceThroughStream(self, stream, text):
        """Speak a ring announcement in the voice of the stream being tuned.

        Adjusting a voice while listening to a different one - different voice, place and
        speed - tells you nothing about what you are changing. The announcement is the
        sample.

        Returns True if it was handled, so the caller knows to silence the main voice.
        """
        voice = self.voiceFor(stream)
        player = self.playerForVoicedStream(stream)
        if voice is None or player is None:
            return False
        try:
            voice.speak(text, player, kind="ui", interrupt=True)
            log.info("Spatial Typing Feedback: announced %r via %s" % (text, stream.value))
            return True
        except Exception:  # noqa: BLE001
            log.error("Spatial Typing Feedback: could not announce change", exc_info=True)
            return False

    def suppressText(self, text):
        """Drop one specific thing NVDA's main voice is about to say.

        Matched on the exact text rather than "the next utterance", so a missed match
        costs one duplicated announcement instead of silently swallowing something
        unrelated. Also time-limited, in the same spirit as NVDA's own typed-character
        suppression.
        """
        self._suppressText = text
        self._suppressUntil = time.time() + 0.5

    def _errorAnnotations(self):
        """NVDA's own wording for the spelling and grammar markers.

        Read from NVDA's translation catalogue rather than ours - builtins._ stays bound
        to NVDA's domain, while addonHandler.initTranslation rebinds _ inside this module
        to the add-on's. Matching on our own translation would fail in every language but
        English.
        """
        if self._annotationTexts is None:
            try:
                import builtins

                translate = builtins._
                self._annotationTexts = {
                    translate("spelling error"),
                    translate("out of spelling error"),
                    translate("grammar error"),
                    translate("out of grammar error"),
                }
            except Exception:  # noqa: BLE001
                log.debugWarning("Could not read NVDA's error wording", exc_info=True)
                self._annotationTexts = set()
        return self._annotationTexts

    def _moveAnnotationsToErrorStream(self, speechSequence):
        """Take "spelling error" and friends out of the sentence and speak them left.

        They describe an error rather than forming part of the text, so they belong with
        the other error feedback. Returns the sequence with them removed.

        Only exact matches are taken, and only when there is somewhere to put them - if
        the annotation voice is unavailable the strings are left alone rather than
        silently dropped.
        """
        if not (self._annotationMode & 2) or not self.shouldSpeak():
            return speechSequence
        # In filter-only mode there is no annotation voice, so borrow the character one.
        # That isolates the filter from the extra-synthesizer question.
        voice = self._annotationVoice or self._voice
        player = (
            self._annotationVoice.player if self._annotationVoice is not None
            else self._charPlayer
        )
        if voice is None or player is None:
            return speechSequence
        annotations = self._errorAnnotations()
        if not annotations:
            return speechSequence
        found = [
            item for item in speechSequence
            if isinstance(item, str) and item.strip() in annotations
        ]
        if not found:
            return speechSequence
        remaining = [item for item in speechSequence if item not in found]
        for text in found:
            try:
                voice.speak(text, player, kind="annotation")
            except Exception:  # noqa: BLE001
                log.error("Spatial Typing Feedback: could not speak annotation", exc_info=True)
                return speechSequence
        return remaining

    def _filterSpeech(self, speechSequence, **kwargs):
        speechSequence = self._moveAnnotationsToErrorStream(speechSequence)
        wanted = self._suppressText
        if not wanted:
            return speechSequence
        if time.time() > (self._suppressUntil or 0):
            self._suppressText = None
            return speechSequence
        spoken = " ".join(item for item in speechSequence if isinstance(item, str)).strip()
        if spoken != wanted.strip():
            log.info(
                "Spatial Typing Feedback: not suppressing %r (waiting for %r)"
                % (spoken, wanted.strip()),
            )
        if spoken == wanted.strip():
            self._suppressText = None
            log.info("Spatial Typing Feedback: suppressed duplicate %r" % spoken)
            return []
        return speechSequence

    def _installSpeechFilter(self):
        try:
            speech.extensions.filter_speechSequence.register(self._filterSpeech)
            self._speechFilterInstalled = True
            log.info("Spatial Typing Feedback: speech filter installed")
        except Exception:  # noqa: BLE001
            log.error(
                "Spatial Typing Feedback: could not hook filter_speechSequence; ring "
                "changes will be announced twice",
                exc_info=True,
            )

    def _removeSpeechFilter(self):
        if not getattr(self, "_speechFilterInstalled", False):
            return
        try:
            speech.extensions.filter_speechSequence.unregister(self._filterSpeech)
        except Exception:  # noqa: BLE001
            log.debugWarning("Error unregistering the speech filter", exc_info=True)
        self._speechFilterInstalled = False

    def getMainRateBoost(self):
        synth = self._synth()
        try:
            return bool(getattr(synth, "rateBoost", False)) if synth else False
        except Exception:  # noqa: BLE001
            return False

    def setMainRateBoost(self, enable):
        synth = self._synth()
        if synth is None:
            return
        try:
            synth.rateBoost = enable
            synth.saveSettings()
        except Exception:  # noqa: BLE001
            log.debugWarning("Could not set main rate boost", exc_info=True)

    def getMainPunctuation(self):
        synth = self._synth()
        try:
            return bool(getattr(synth, "punctuationSilence", True)) if synth else True
        except Exception:  # noqa: BLE001
            return True

    def setMainPunctuation(self, enable):
        synth = self._synth()
        if synth is None:
            return
        try:
            synth.punctuationSilence = enable
            synth.saveSettings()
        except Exception:  # noqa: BLE001
            log.debugWarning("Could not set main punctuation silence", exc_info=True)

    def mainSupportsPunctuation(self):
        synth = self._synth()
        if synth is None:
            return False
        try:
            return bool(getattr(synth, "supportsPunctuationSilence", False))
        except Exception:  # noqa: BLE001
            return False

    def getMainPitch(self):
        synth = self._synth()
        try:
            return int(getattr(synth, "pitch", 50)) if synth else 50
        except Exception:  # noqa: BLE001
            return 50

    def setMainPitch(self, value):
        synth = self._synth()
        if synth is None:
            return
        try:
            synth.pitch = value
            synth.saveSettings()
        except Exception:  # noqa: BLE001
            log.debugWarning("Could not set main pitch", exc_info=True)

    # -- saving when you leave the ring ------------------------------------------

    #: Ring scripts all live in globalCommands and are named ...SynthSetting. Matching on
    #: the name rather than on key combinations means user-rebound keys still work.
    RING_SCRIPT_MARKER = "SynthSetting"

    def _isRingGesture(self, gesture):
        try:
            # gesture.script caches its own lookup; NVDA reads the same property a moment
            # later. This hook runs on every keystroke, so avoid resolving twice.
            script = getattr(gesture, "script", None)
            if script is None:
                script = scriptHandler.findScript(gesture)
        except Exception:  # noqa: BLE001
            return False
        return self.RING_SCRIPT_MARKER in (getattr(script, "__name__", "") or "")

    def _onGesture(self, gesture, **kwargs):
        """Save settings on the way out of the ring.

        The moment a key arrives that is not a ring command, tuning is over - so commit
        it. That transition is the only trigger; there is no timer deciding on the
        user's behalf when they have finished.

        Never blocks a gesture: this only observes.
        """
        try:
            if _config.isDirty() and not self._isRingGesture(gesture):
                _config.flush()
        except Exception:  # noqa: BLE001
            log.debugWarning("Error saving on leaving the ring", exc_info=True)
        return True

    def _installGestureHook(self):
        try:
            inputCore.decide_executeGesture.register(self._onGesture)
            self._gestureHookInstalled = True
        except Exception:  # noqa: BLE001
            log.error(
                "Spatial Typing Feedback: could not hook gestures; settings will only "
                "save when the add-on unloads or NVDA exits",
                exc_info=True,
            )

    def _removeGestureHook(self):
        if not getattr(self, "_gestureHookInstalled", False):
            return
        try:
            inputCore.decide_executeGesture.unregister(self._onGesture)
        except Exception:  # noqa: BLE001
            log.debugWarning("Error unregistering the gesture hook", exc_info=True)
        self._gestureHookInstalled = False

    def _stop(self):
        _config.flush()
        self._removeGestureHook()
        self._removeSpeechFilter()
        _ring.uninstall(self._originalRing)
        self._originalRing = None
        self._removeCancelHook()
        if self._echo is not None:
            self._echo.uninstall()
            self._echo = None
        if self._errors is not None:
            self._errors.uninstall()
            self._errors = None
        if self._voice is not None:
            self._voice.terminate()
            self._voice = None
        if self._annotationVoice is not None:
            self._annotationVoice.terminate()
            self._annotationVoice = None
        for player in (self._charPlayer, self._errorPlayer):
            if player is not None:
                player.terminate()
        self._charPlayer = None
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

    def shouldSpeak(self):
        """Whether our voices may speak right now.

        Our audio bypasses NVDA's speak(), so nothing stops us automatically. Silencing
        the screen reader has to silence us too - anything else is both wrong and
        confusing, since a quiet main voice next to a chattering echo looks exactly like
        this add-on being broken.

        Covers speech mode (off, beeps, on-demand) and sleep mode, which is how NVDA is
        told to stay out of the way in self-voicing applications.
        """
        try:
            mode = speech.getState().speechMode
            if mode in (speech.SpeechMode.off, speech.SpeechMode.beeps):
                return False
            if mode == speech.SpeechMode.onDemand:
                # On demand means speak only when explicitly asked. Typing echo is not
                # an explicit request.
                return False
        except Exception:  # noqa: BLE001
            pass
        try:
            focus = api.getFocusObject()
            if focus is not None and focus.sleepMode:
                return False
        except Exception:  # noqa: BLE001
            pass
        return True

    def _speakChar(self, text):
        if self._voice is None or not self.shouldSpeak():
            return
        # Interrupt, matching NVDA's own speechInterruptForCharacters. Synthesis takes
        # longer than a fast typist takes to reach the next key, so queueing every
        # character makes the echo drift further and further behind the keyboard.
        self._voice.speak(text, self._charPlayer, kind="char", interrupt=True)

    # -- following NVDA's own speech state ---------------------------------------

    def _onSpeechCanceled(self):
        """Stop our streams whenever NVDA's speech is stopped.

        Control, or anything else that calls cancelSpeech, silences the main voice. Our
        voices are separate players NVDA knows nothing about, so without this they carry
        on talking over the silence - and the user has no way to shut them up.
        """
        try:
            for voice in (self._voice, self._annotationVoice):
                if voice is not None:
                    voice.cancel()
            for player in (self._charPlayer, self._errorPlayer):
                if player is not None:
                    player.stop()
        except Exception:  # noqa: BLE001
            log.error("Spatial Typing Feedback: error cancelling speech", exc_info=True)

    def _installCancelHook(self):
        try:
            speech.extensions.speechCanceled.register(self._onSpeechCanceled)
            self._cancelHookInstalled = True
            log.info("Spatial Typing Feedback: speech cancel hook installed")
        except Exception:  # noqa: BLE001
            log.error(
                "Spatial Typing Feedback: could not hook speechCanceled; Control will not "
                "stop the secondary voices",
                exc_info=True,
            )

    def _removeCancelHook(self):
        if not getattr(self, "_cancelHookInstalled", False):
            return
        try:
            speech.extensions.speechCanceled.unregister(self._onSpeechCanceled)
        except Exception:  # noqa: BLE001
            log.debugWarning("Error unregistering speechCanceled", exc_info=True)
        self._cancelHookInstalled = False

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
        if self._errorPlayer is not None:
            errorWav = os.path.join(globalVars.appDir, "waves", "textError.wav")
            if os.path.isfile(errorWav):
                self._errorPlayer.feedWaveFile(errorWav)
        if self._voice is None:
            # Say why, rather than leaving a buzz and silence to be interpreted.
            ui.message(
                self._voiceError
                # Translators: Reported when the secondary voice is not running.
                or _("Error alert only. The secondary voice is not running"),
            )
            return
        if self.engineName:
            # ui.message goes through NVDA's main voice, which is the centre stream -
            # so this doubles as the test for it.
            # Translators: Reported by the test command; {engine} is e.g. "OneCore".
            ui.message(_("Main voice, centre. Using {engine}").format(engine=self.engineName))
        # Translators: Spoken by the test command, from the character stream.
        self._voice.speak(_("characters, right"), self._charPlayer, kind="test")
