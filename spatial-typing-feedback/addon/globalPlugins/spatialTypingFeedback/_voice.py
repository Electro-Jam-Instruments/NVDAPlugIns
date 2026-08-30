# A second, independent Windows OneCore voice.
#
# See docs/architecture-decisions.md Decision D3.
#
# OneCore is token-based: ocSpeech_initialize() returns a HANDLE and every subsequent
# call takes it. Instances are therefore independent, unlike eSpeak which is bound to
# module-level global state. We go to the ocSpeech layer directly rather than
# instantiating OneCoreSynthDriver, because that way we own the callback - and
# therefore the WavePlayer we need to pan.

import ctypes
import io
import wave
from ctypes.wintypes import HANDLE

import comtypes
import NVDAHelper
from logHandler import log

from ._audio import WAVE_HEADER_LENGTH, PannedPlayer

#: Signature of the callback ocSpeech hands audio back through.
ocSpeech_Callback = ctypes.CFUNCTYPE(None, ctypes.c_void_p, ctypes.c_int, ctypes.c_wchar_p)

#: Rate mapping, matching NVDA's oneCore driver exactly.
#: Rate boost is nothing more than raising the top of this range.
MIN_RATE = 0.5
DEFAULT_MAX_RATE = 1.5
BOOSTED_MAX_RATE = 6.0

MIN_PITCH = 0.0
MAX_PITCH = 2.0


def _percentToParam(percent, minVal, maxVal):
    return float(percent) / 100 * (maxVal - minVal) + minVal


class OneCoreVoice(object):
    """One independent OneCore instance, rendered to a panned player of our own.

    Each instance owns a token, a callback and a PannedPlayer, so two of them can speak
    at different positions and levels at the same instant - which is required, because
    one space fires the word echo and the character echo together.
    """

    def __init__(self, name="secondary"):
        self.name = name
        self.available = False
        self.supportsProsodyOptions = False
        self._dll = None
        self._token = None
        self._callbackInst = None
        self._earlyExitCB = False
        self._queue = []
        self._rate = 50
        self._rateBoost = False
        self._pitch = 50
        self._volume = 100
        self.player = PannedPlayer()

    # -- lifecycle --------------------------------------------------------------

    def initialize(self):
        """Create our own ocSpeech token. Returns True on success.

        This is the assumption the whole design rests on: that ocSpeech_initialize can
        be called a second time while NVDA's own OneCore driver holds a live token.
        """
        try:
            self._dll = NVDAHelper.getHelperLocalWin10Dll()
        except Exception:  # noqa: BLE001
            log.error("Spatial Typing Feedback: could not load helperLocalWin10", exc_info=True)
            return False
        try:
            self._dll.ocSpeech_initialize.restype = HANDLE
            self._dll.ocSpeech_getCurrentVoiceLanguage.restype = ctypes.c_wchar_p
            self._dll.ocSpeech_supportsProsodyOptions.restype = ctypes.c_bool
            self.supportsProsodyOptions = bool(self._dll.ocSpeech_supportsProsodyOptions())
            if self.supportsProsodyOptions:
                self._dll.ocSpeech_getPitch.restype = ctypes.c_double
                self._dll.ocSpeech_getVolume.restype = ctypes.c_double
                self._dll.ocSpeech_getRate.restype = ctypes.c_double
            else:
                log.warning(
                    "Spatial Typing Feedback: OneCore prosody options unsupported; "
                    "rate, pitch and volume cannot be set on the secondary voice",
                )
            self._callbackInst = ocSpeech_Callback(self._callback)
            token = HANDLE()
            token.value = self._dll.ocSpeech_initialize(self._callbackInst)
            if not token.value:
                log.error(
                    "Spatial Typing Feedback: ocSpeech_initialize returned a null token "
                    "for the %s voice" % self.name,
                )
                return False
            self._token = token
            self._dll.ocSpeech_getVoices.restype = comtypes.BSTR
            self._dll.ocSpeech_getCurrentVoiceId.restype = ctypes.c_wchar_p
        except Exception:  # noqa: BLE001
            log.error(
                "Spatial Typing Feedback: failed to create a second OneCore instance (%s)"
                % self.name,
                exc_info=True,
            )
            self._token = None
            self._callbackInst = None
            return False
        self.available = True
        log.info("Spatial Typing Feedback: %s OneCore voice initialized" % self.name)
        return True

    def terminate(self):
        # Stop pending callbacks touching us, exactly as NVDA's driver does.
        self._earlyExitCB = True
        self._queue = []
        self.player.terminate()
        if self._token is not None and self._dll is not None:
            try:
                self._dll.ocSpeech_terminate(self._token)
            except Exception:  # noqa: BLE001
                log.debugWarning("Error terminating ocSpeech token", exc_info=True)
        # Drop the ctypes callback instance; it holds a reference to a bound method.
        self._token = None
        self._callbackInst = None
        self.available = False

    # -- voice parameters -------------------------------------------------------

    def getAvailableVoiceIds(self):
        """Return a list of (id, onecoreIndex) for every voice OneCore reports."""
        if not self.available:
            return []
        try:
            voicesStr = self._dll.ocSpeech_getVoices(self._token).split("|")
        except Exception:  # noqa: BLE001
            log.debugWarning("Error fetching OneCore voices", exc_info=True)
            return []
        result = []
        for index, voiceStr in enumerate(voicesStr):
            parts = voiceStr.split(":")
            if len(parts) < 3:
                continue
            result.append((parts[0], index))
        return result

    def setVoiceById(self, voiceId):
        """Match the main synth's voice. Returns True if found and set."""
        if not self.available or not voiceId:
            return False
        for vid, index in self.getAvailableVoiceIds():
            if vid == voiceId:
                try:
                    self._dll.ocSpeech_setVoice(self._token, index)
                    return True
                except Exception:  # noqa: BLE001
                    log.debugWarning("Error setting OneCore voice", exc_info=True)
                    return False
        log.debugWarning("Spatial Typing Feedback: voice %r not found for %s" % (voiceId, self.name))
        return False

    def setRate(self, rate, rateBoost=False):
        """Set rate 0-100, honouring rate boost.

        Rate boost is only a range change: 0-100 maps onto 0.5-6.0 instead of 0.5-1.5.
        The user is on OneCore for this, so the secondary voice has to have it too -
        an echo capped at normal speed would still be arriving after the main voice had
        moved on.
        """
        self._rate = rate
        self._rateBoost = rateBoost
        if not self.supportsProsodyOptions:
            return
        maxRate = BOOSTED_MAX_RATE if rateBoost else DEFAULT_MAX_RATE
        self._queueParam(self._dll.ocSpeech_setRate, _percentToParam(rate, MIN_RATE, maxRate))

    def setPitch(self, pitch):
        self._pitch = pitch
        if not self.supportsProsodyOptions:
            return
        self._queueParam(self._dll.ocSpeech_setPitch, _percentToParam(pitch, MIN_PITCH, MAX_PITCH))

    def setVolume(self, volume):
        """Synth volume 0-100. Distinct from the stream level, which is a channel gain."""
        self._volume = volume
        if not self.supportsProsodyOptions:
            return
        self._queueParam(self._dll.ocSpeech_setVolume, volume / 100.0)

    def setPosition(self, pan, volumePercent=100):
        self.player.setPosition(pan, volumePercent)

    # -- speaking ---------------------------------------------------------------

    def speak(self, text):
        if not self.available or not text:
            return
        self._queue.append(text)
        self._processQueue()

    def cancel(self):
        # Keep queued parameter changes, drop queued text - same as NVDA's driver.
        self._queue = [item for item in self._queue if not isinstance(item, str)]
        self.player.stop()

    def _queueParam(self, func, value):
        self._queue.append((func, value))
        self._processQueue()

    def _processQueue(self):
        if not self.available:
            return
        while self._queue:
            item = self._queue.pop(0)
            if isinstance(item, tuple):
                func, value = item
                try:
                    func(self._token, ctypes.c_double(value))
                except Exception:  # noqa: BLE001
                    log.debugWarning("Error applying OneCore parameter", exc_info=True)
                continue
            try:
                # Async: _callback fires on a background thread when audio is ready,
                # and processes the queue again from there.
                self._dll.ocSpeech_speak(self._token, item)
            except Exception:  # noqa: BLE001
                log.error("Spatial Typing Feedback: ocSpeech_speak failed", exc_info=True)
                continue
            return

    def _callback(self, bytesPtr, length, markers):
        """Audio has been rendered. Runs on a background thread."""
        if self._earlyExitCB:
            return
        try:
            if length == 0:
                log.debugWarning("Spatial Typing Feedback: %s voice produced no audio" % self.name)
                return
            header = ctypes.string_at(bytesPtr, WAVE_HEADER_LENGTH)
            with wave.open(io.BytesIO(header), "r") as wav:
                samplesPerSec = wav.getframerate()
                channels = wav.getnchannels()
                sampleWidth = wav.getsampwidth()
                dataLen = wav.getnframes() * channels * sampleWidth
            data = ctypes.string_at(bytesPtr + WAVE_HEADER_LENGTH, dataLen)
            if channels == 1:
                self.player.feedMono(data, samplesPerSec)
            else:
                self.player.feedStereo(data, samplesPerSec)
        except Exception:  # noqa: BLE001
            log.error("Spatial Typing Feedback: error in ocSpeech callback", exc_info=True)
        finally:
            if not self._earlyExitCB:
                self._processQueue()
