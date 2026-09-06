# The SAPI 5 fallback voice.
#
# The primary engine is _winrt.py, which activates a Windows SpeechSynthesizer directly
# and offers the same voices as NVDA's own synthesizer.
#
# This is for machines where that fails: no Windows voices installed, or an activation
# error. It uses older voice builds and tops out near three times normal speed, so it is a
# genuine downgrade - but a working downgrade beats no character echo.
#
# Its interface is deliberately identical to WinRTVoice. A fallback that cannot be called
# the same way as the thing it replaces is not a fallback; letting the two drift apart once
# cost an evening of silent keystrokes.

import math
import queue
import threading

import comtypes
import comtypes.client
from logHandler import log

from ._audio import PannedPlayer

#: SpeechAudioFormatType.SAFT22kHz16BitMono
SAFT_22KHZ_16BIT_MONO = 22
SAMPLES_PER_SEC = 22050

#: SAPI 5 rate runs -10..+10, where +10 is roughly three times normal speed.
SAPI_RATE_MIN = -10
SAPI_RATE_MAX = 10
SAPI_RATE_MAX_MULTIPLIER = 3.0

#: OneCore's rate mapping, so we can match the speed the user actually has.
ONECORE_MIN_RATE = 0.5
ONECORE_DEFAULT_MAX_RATE = 1.5
ONECORE_BOOSTED_MAX_RATE = 6.0

_STOP = object()


def oneCoreRateToSapi(rate, rateBoost):
    """Convert an NVDA OneCore rate into the nearest SAPI 5 rate.

    OneCore's raw rate *is* a speed multiplier: 0-100 maps onto 0.5-1.5, or 0.5-6.0
    with rate boost. SAPI's scale is roughly logarithmic, about 3x at +10. Converting
    through the multiplier keeps the secondary voice at the speed the user actually
    has, rather than at whatever number the two scales happen to share.

    Rate boost is why OneCore was chosen in the first place, so an echo that lags behind
    the main voice would defeat the point.
    """
    maxRate = ONECORE_BOOSTED_MAX_RATE if rateBoost else ONECORE_DEFAULT_MAX_RATE
    multiplier = float(rate) / 100 * (maxRate - ONECORE_MIN_RATE) + ONECORE_MIN_RATE
    multiplier = max(0.1, multiplier)
    sapiRate = SAPI_RATE_MAX * math.log(multiplier) / math.log(SAPI_RATE_MAX_MULTIPLIER)
    return int(round(max(SAPI_RATE_MIN, min(SAPI_RATE_MAX, sapiRate))))


def voiceKeywordFromOneCoreId(voiceId):
    """Pull a matchable name out of a OneCore voice ID.

    OneCore IDs look like
    ...\\Speech_OneCore\\Voices\\Tokens\\MSTTS_V110_enUS_ZiraM
    and the nearest SAPI voice is described as "Microsoft Zira Desktop - English".
    The shared part is the name, so that is what we match on.
    """
    if not voiceId:
        return None
    token = voiceId.rstrip("\\").split("\\")[-1]
    parts = token.split("_")
    if not parts:
        return None
    name = parts[-1]
    # A trailing capital denotes the OneCore build (ZiraM, DavidM); drop it.
    if len(name) > 2 and name[-1].isupper() and name[-2].islower():
        name = name[:-1]
    return name or None


class SapiVoice(object):
    """A SAPI 5 voice rendering into panned players we own.

    Synthesis runs on a worker thread. SAPI's Speak() into a memory stream is
    synchronous, and blocking NVDA's main thread on every typed character would be
    worse than no echo at all.

    Each queued utterance carries its destination player, so one voice serves several
    streams at different positions and levels.
    """

    def __init__(self, name="secondary"):
        self.name = name
        self.available = False
        self.spokeSuccessfully = False
        self.lastError = None
        self._voice = None
        self._format = None
        self._queue = None
        self._thread = None
        self._stopping = False
        self.player = PannedPlayer(name)

    # -- lifecycle --------------------------------------------------------------

    def initialize(self):
        try:
            self._voice = comtypes.client.CreateObject("SAPI.SpVoice")
            self._format = comtypes.client.CreateObject("SAPI.SpAudioFormat")
            self._format.Type = SAFT_22KHZ_16BIT_MONO
        except Exception as e:  # noqa: BLE001
            self.lastError = str(e)
            log.error(
                "Spatial Typing Feedback: could not create a SAPI 5 voice (%s)" % self.name,
                exc_info=True,
            )
            self._voice = None
            return False
        self._queue = queue.Queue()
        self._stopping = False
        self._thread = threading.Thread(
            target=self._worker,
            name="spatialTypingFeedback-%s" % self.name,
            daemon=True,
        )
        self._thread.start()
        self.available = True
        log.info("Spatial Typing Feedback: %s SAPI voice initialized" % self.name)
        return True

    def terminate(self):
        self._stopping = True
        self.available = False
        if self._queue is not None:
            try:
                self._queue.put_nowait(_STOP)
            except Exception:  # noqa: BLE001
                pass
        if self._thread is not None and self._thread.is_alive():
            self._thread.join(timeout=2.0)
        self._thread = None
        self._queue = None
        self._voice = None
        self._format = None
        self.player.terminate()

    # -- voice parameters -------------------------------------------------------

    def _getAvailableVoiceNames(self):
        if self._voice is None:
            return []
        try:
            tokens = self._voice.GetVoices()
            return [tokens.Item(i).GetDescription() for i in range(tokens.Count)]
        except Exception:  # noqa: BLE001
            log.debugWarning("Error listing SAPI voices", exc_info=True)
            return []

    def matchVoice(self, voiceId):
        """Pick the SAPI voice closest to the main synth's voice.

        Separation is meant to be positional, not timbral, so the nearest available
        match beats an arbitrary default.
        """
        if self._voice is None or not voiceId:
            return False
        # Exact match first: the caller may be handing back one of our own descriptions.
        try:
            tokens = self._voice.GetVoices()
            for i in range(tokens.Count):
                token = tokens.Item(i)
                if token.GetDescription() == voiceId:
                    self._voice.Voice = token
                    return True
        except Exception:  # noqa: BLE001
            log.debugWarning("Error matching SAPI voice by description", exc_info=True)
        keyword = voiceKeywordFromOneCoreId(voiceId)
        if not keyword:
            return False
        try:
            tokens = self._voice.GetVoices()
            for i in range(tokens.Count):
                token = tokens.Item(i)
                if keyword.lower() in token.GetDescription().lower():
                    self._voice.Voice = token
                    log.info(
                        "Spatial Typing Feedback: %s matched to SAPI voice %r"
                        % (self.name, token.GetDescription()),
                    )
                    return True
        except Exception:  # noqa: BLE001
            log.debugWarning("Error matching SAPI voice", exc_info=True)
            return False
        log.info(
            "Spatial Typing Feedback: no SAPI voice matching %r; using the default. "
            "Available: %s" % (keyword, self._getAvailableVoiceNames()),
        )
        return False

    def setRate(self, rate, rateBoost=False):
        if self._voice is None:
            return
        sapiRate = oneCoreRateToSapi(rate, rateBoost)
        try:
            self._voice.Rate = sapiRate
            log.info(
                "Spatial Typing Feedback: %s rate %s (boost=%s) -> SAPI rate %s"
                % (self.name, rate, rateBoost, sapiRate),
            )
        except Exception:  # noqa: BLE001
            log.debugWarning("Error setting SAPI rate", exc_info=True)

    def setVolume(self, volume):
        """Synth volume 0-100. Distinct from the stream level, which is a channel gain."""
        if self._voice is None:
            return
        try:
            self._voice.Volume = max(0, min(100, int(volume)))
        except Exception:  # noqa: BLE001
            log.debugWarning("Error setting SAPI volume", exc_info=True)

    def setPitch(self, pitch):
        # SAPI pitch is per-utterance XML rather than a property, not a settable one.
        pass

    # -- interface parity with WinRTVoice -----------------------------------------

    def setRatePercent(self, percent, rateBoost=True):
        """Set speed from a 0-100 value.

        OneCore maps this onto 0.5x-1.5x, or 0.5x-6.0x with rate boost; SAPI's scale is
        -10..+10 and roughly logarithmic. Convert through the speed multiplier so the
        two engines land at a comparable speed rather than a comparable number.
        """
        percent = max(0, min(100, int(percent)))
        maxMultiplier = ONECORE_BOOSTED_MAX_RATE if rateBoost else ONECORE_DEFAULT_MAX_RATE
        multiplier = ONECORE_MIN_RATE + (percent / 100.0) * (
            maxMultiplier - ONECORE_MIN_RATE
        )
        sapiRate = SAPI_RATE_MAX * math.log(multiplier) / math.log(SAPI_RATE_MAX_MULTIPLIER)
        sapiRate = int(round(max(SAPI_RATE_MIN, min(SAPI_RATE_MAX, sapiRate))))
        if self._voice is None:
            return False
        try:
            self._voice.Rate = sapiRate
            log.info(
                "Spatial Typing Feedback: %s speed %d%% -> %.2fx -> SAPI rate %d"
                % (self.name, percent, multiplier, sapiRate),
            )
            return True
        except Exception:  # noqa: BLE001
            log.debugWarning("Could not set SAPI rate", exc_info=True)
            return False

    def setPitchPercent(self, percent):
        # Not settable as a property on SAPI. Accepted so the interfaces match.
        return False

    def setPunctuationSilence(self, enable):
        # No SAPI equivalent. Accepted so the interfaces match.
        return False

    def getAvailableVoiceIds(self):
        """[(id, index)], matching WinRTVoice.

        SAPI has no OneCore-style IDs, so the description doubles as the identifier.
        """
        if self._voice is None:
            return []
        try:
            tokens = self._voice.GetVoices()
            return [(tokens.Item(i).GetDescription(), i) for i in range(tokens.Count)]
        except Exception:  # noqa: BLE001
            log.debugWarning("Error listing SAPI voices", exc_info=True)
            return []

    def getVoiceDisplayName(self, voiceId):
        """Matching WinRTVoice. SAPI has no separate ID, so the name is the ID."""
        return voiceId or None

    # -- speaking ---------------------------------------------------------------

    def speak(self, text, player=None, kind=None, interrupt=False):
        """Queue text for a destination player.

        Signature deliberately identical to WinRTVoice.speak. This class exists to
        stand in for that one, and a substitute that cannot be called the same way is
        not a substitute.

        @param kind: a label such as "char" or "word", for selective interruption.
        @param interrupt: drop anything pending OF THE SAME KIND first, so characters
            interrupt characters without cancelling a word queued behind them.
        """
        if not self.available or not text or self._queue is None:
            return
        if interrupt:
            self.cancelKind(kind)
        self._queue.put((text, player if player is not None else self.player, kind))

    def cancelKind(self, kind):
        """Drop pending utterances of one kind, leaving the others alone."""
        if self._queue is None:
            return
        keep = []
        try:
            while True:
                item = self._queue.get_nowait()
                if item is _STOP:
                    keep.append(item)
                elif item[2] != kind:
                    keep.append(item)
                else:
                    item[1].stop()
        except queue.Empty:
            pass
        for item in keep:
            self._queue.put(item)

    def cancel(self):
        if self._queue is None:
            return
        try:
            while True:
                item = self._queue.get_nowait()
                if item is not _STOP:
                    item[1].stop()
        except queue.Empty:
            pass
        self.player.stop()

    def _worker(self):
        try:
            comtypes.CoInitializeEx()
        except Exception:  # noqa: BLE001
            log.debugWarning("CoInitializeEx failed on the synthesis thread", exc_info=True)
        while not self._stopping:
            try:
                item = self._queue.get()
            except Exception:  # noqa: BLE001
                break
            if item is _STOP or self._stopping:
                break
            text, player = item[0], item[1]
            try:
                data = self._synthesize(text)
            except Exception:  # noqa: BLE001
                log.error("Spatial Typing Feedback: synthesis failed", exc_info=True)
                continue
            if not data:
                log.debugWarning("Spatial Typing Feedback: no audio produced for %r" % text)
                continue
            self.spokeSuccessfully = True
            player.feedMono(data, SAMPLES_PER_SEC)

    def _synthesize(self, text):
        """Render text to raw 16-bit mono PCM.

        A fresh memory stream per utterance: reusing one appends each render to the
        last, so the echo would replay everything typed so far.
        """
        stream = comtypes.client.CreateObject("SAPI.SpMemoryStream")
        stream.Format = self._format
        self._voice.AudioOutputStream = stream
        # 0 = SVSFDefault, i.e. synchronous. We are on a worker thread, so that is fine.
        self._voice.Speak(text, 0)
        return bytes(bytearray(stream.GetData()))


#: The plugin talks to whatever the secondary voice happens to be, so the engine can
#: change (D12) without touching the rest of the add-on.
SecondaryVoice = SapiVoice
