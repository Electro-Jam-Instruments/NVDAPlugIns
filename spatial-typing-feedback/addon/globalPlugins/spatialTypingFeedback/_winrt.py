# The character voice: a Windows SpeechSynthesizer we activate ourselves.
#
# The same engine NVDA's OneCore synthesizer uses - the same class, the same voices, the
# same voice IDs - but activated directly rather than through NVDA.
#
# That is necessary because nvdaHelper's ocSpeech wrapper is a process-wide singleton
# holding one synthesizer, and NVDA's own driver occupies it. The limit is NVDA's wrapper,
# not Windows: Windows.Media.SpeechSynthesis.SpeechSynthesizer is a public, freely
# instantiable WinRT class. Going straight to it means we depend on nothing of NVDA's for
# the voice.
#
# ON THE VTABLE CALLS
#
# WinRT generic interfaces such as IAsyncOperation<T> have IIDs computed by hashing their
# type arguments, which is impractical here. Their vtable layouts are fixed, so methods
# are called by slot instead. Every IID and slot below was read off the running system via
# IInspectable::GetIids and confirmed with GetRuntimeClassName, not guessed - one of them
# differs from the value commonly quoted.

import ctypes
import io
import queue
import threading
import wave

from logHandler import log

from ._audio import PannedPlayer

combase = ctypes.WinDLL("combase.dll")

S_OK = 0
RO_INIT_MULTITHREADED = 1
ASYNC_STARTED, ASYNC_COMPLETED = 0, 1

# Speed is set through SSML prosody, and these limits were MEASURED on a running system
# rather than taken from documentation, because both differ from what you would expect:
#
#   * SpeechSynthesizerOptions.SpeakingRate stores and reads back correctly but has no
#     effect whatsoever on the synthesised audio. Identical byte counts at 0.5x, 1.0x and
#     3.0x, with unique text each time to rule out caching.
#   * SSML prosody rate works, but SATURATES AT 200%. 300%, 500% and 1000% all produce
#     exactly the same audio as 200%. Sending a larger number silently pins the voice at
#     maximum, which sounds like a bug in the mapping.
#
# Measured against rate="100%": 10% -> 0.58x, 100% -> 1.00x, 200% -> 1.54x, and flat
# thereafter.
PROSODY_RATE_MIN = 10
PROSODY_RATE_MAX = 200
#: Without rate boost, a narrower band around normal speed, for fine control.
PROSODY_RATE_PLAIN_MIN = 50
PROSODY_RATE_PLAIN_MAX = 120

CLASS_SYNTH = "Windows.Media.SpeechSynthesis.SpeechSynthesizer"
CLASS_BUFFER = "Windows.Storage.Streams.Buffer"

IID_IInspectable = "{AF86E2E0-B12D-4C6A-9C5A-D7AA65101E90}"
IID_IAsyncInfo = "{00000036-0000-0000-C000-000000000046}"
IID_IRandomAccessStream = "{905A0FE1-BC53-11DF-8C49-001E4FC686DA}"
IID_IInputStream = "{905A0FE2-BC53-11DF-8C49-001E4FC686DA}"
IID_IBufferByteAccess = "{905A0FEF-BC53-11DF-8C49-001E4FC686DA}"
IID_IBufferFactory = "{71AF914D-C10F-484B-BC50-14BC623B3A27}"
IID_ISpeechSynthesizer2 = "{A7C5ECB2-4339-4D6A-BBF8-C7A4F1544C2E}"
IID_ISpeechSynthesizerStatics = "{7D526ECC-7533-4C3F-85BE-888C2BAEEBDC}"
IID_ISpeechSynthesizerOptions2 = "{1CBEF60E-119C-4BED-B118-D250C3A25793}"
IID_ISpeechSynthesizerOptions3 = "{401ED877-902C-4814-A582-A5D0C0769FA8}"
IID_IVoiceInformation = "{B127D6A4-1291-4604-AA9C-83134083352C}"

# Vtable slots. IUnknown occupies 0-2 and IInspectable 3-5 on every WinRT interface.
SLOT_SYNTH_TEXT_ASYNC = 6
SLOT_SYNTH_SSML_ASYNC = 7
SLOT_SYNTH_PUT_VOICE = 8
SLOT_SYNTH_GET_VOICE = 9
SLOT_SYNTH2_GET_OPTIONS = 6
SLOT_OPT2_GET_PITCH, SLOT_OPT2_PUT_PITCH = 6, 7
SLOT_OPT2_GET_VOLUME, SLOT_OPT2_PUT_VOLUME = 8, 9
SLOT_OPT2_GET_RATE, SLOT_OPT2_PUT_RATE = 10, 11
SLOT_OPT3_PUT_PUNCTUATION = 9
SLOT_STATICS_ALL_VOICES = 6
SLOT_VECTOR_GET_AT, SLOT_VECTOR_GET_SIZE = 6, 7
SLOT_VOICE_DISPLAY_NAME, SLOT_VOICE_ID, SLOT_VOICE_LANGUAGE = 6, 7, 8
SLOT_ASYNCINFO_STATUS = 7
SLOT_ASYNCINFO_CANCEL = 9
#: IAsyncOperation has GetResults at 8; IAsyncOperationWithProgress (what ReadAsync
#: returns) puts Progress first, so its GetResults is at 10.
SLOT_OP_GET_RESULTS = 8
SLOT_OP_PROGRESS_GET_RESULTS = 10
SLOT_INPUTSTREAM_READ_ASYNC = 6
SLOT_STREAM_GET_SIZE = 6
SLOT_BUFFER_GET_LENGTH = 7
SLOT_BUFFERFACTORY_CREATE = 6
SLOT_BYTEACCESS_BUFFER = 3

_STOP = object()


class HSTRING(ctypes.c_void_p):
    pass


class GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", ctypes.c_uint32),
        ("Data2", ctypes.c_uint16),
        ("Data3", ctypes.c_uint16),
        ("Data4", ctypes.c_ubyte * 8),
    ]


def _guid(text):
    text = text.strip("{}")
    a, b, c, d, e = text.split("-")
    g = GUID()
    g.Data1, g.Data2, g.Data3 = int(a, 16), int(b, 16), int(c, 16)
    for i, byte in enumerate(bytes.fromhex(d + e)):
        g.Data4[i] = byte
    return g


combase.RoInitialize.argtypes = [ctypes.c_int]
combase.RoInitialize.restype = ctypes.HRESULT
combase.RoUninitialize.restype = None
combase.WindowsCreateString.argtypes = [ctypes.c_wchar_p, ctypes.c_uint32, ctypes.POINTER(HSTRING)]
combase.WindowsCreateString.restype = ctypes.HRESULT
combase.WindowsDeleteString.argtypes = [HSTRING]
combase.WindowsGetStringRawBuffer.argtypes = [HSTRING, ctypes.POINTER(ctypes.c_uint32)]
combase.WindowsGetStringRawBuffer.restype = ctypes.c_wchar_p
combase.RoActivateInstance.argtypes = [HSTRING, ctypes.POINTER(ctypes.c_void_p)]
combase.RoActivateInstance.restype = ctypes.HRESULT
combase.RoGetActivationFactory.argtypes = [
    HSTRING, ctypes.POINTER(GUID), ctypes.POINTER(ctypes.c_void_p),
]
combase.RoGetActivationFactory.restype = ctypes.HRESULT


class _HString(object):
    """An HSTRING with an explicit lifetime.

    Usable as a context manager for SYNCHRONOUS calls only. An asynchronous call keeps
    reading the string after it returns, so freeing it at the end of the `with` block is
    a use-after-free inside the speech engine - which crashes the process with
    STATUS_STACK_BUFFER_OVERRUN, not with anything that looks like our fault. For those,
    hold the object and call free() once the operation has completed.
    """

    def __init__(self, text):
        self.handle = HSTRING()
        combase.WindowsCreateString(text, len(text), ctypes.byref(self.handle))

    def free(self):
        if self.handle:
            combase.WindowsDeleteString(self.handle)
            self.handle = HSTRING()

    def __enter__(self):
        return self.handle

    def __exit__(self, *exc):
        self.free()


def _fromHString(handle):
    if not handle:
        return None
    length = ctypes.c_uint32()
    return combase.WindowsGetStringRawBuffer(handle, ctypes.byref(length))


def _call(ptr, slot, argtypes, *args):
    vtbl = ctypes.cast(ptr, ctypes.POINTER(ctypes.c_void_p)).contents.value
    fn = ctypes.cast(vtbl, ctypes.POINTER(ctypes.c_void_p))[slot]
    return ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p, *argtypes)(fn)(ptr, *args)


def _qi(ptr, iidText):
    iid = _guid(iidText)
    out = ctypes.c_void_p()
    hr = _call(ptr, 0, [ctypes.c_void_p, ctypes.POINTER(ctypes.c_void_p)],
               ctypes.byref(iid), ctypes.byref(out))
    return out if hr == S_OK and out.value else None


def _release(ptr):
    if ptr and getattr(ptr, "value", None):
        try:
            vtbl = ctypes.cast(ptr, ctypes.POINTER(ctypes.c_void_p)).contents.value
            fn = ctypes.cast(vtbl, ctypes.POINTER(ctypes.c_void_p))[2]
            ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)(fn)(ptr)
        except Exception:  # noqa: BLE001
            pass


def _getString(ptr, slot):
    handle = HSTRING()
    if _call(ptr, slot, [ctypes.POINTER(HSTRING)], ctypes.byref(handle)) != S_OK:
        return None
    try:
        return _fromHString(handle)
    finally:
        combase.WindowsDeleteString(handle)


def _await(op, resultsSlot, watchdog=10.0):
    """Wait for an async operation to STOP, then take its result.

    Returns (result, stopped). `stopped` is False only when the operation could not be
    proven to have finished - in which case the caller must not free anything the
    operation may still be touching, including its arguments.

    There is deliberately no "give up and carry on" path. Abandoning a running operation
    and then releasing its buffers is how you corrupt the speech engine's heap, and it
    surfaces as the whole process dying inside MSTTSEngine_OneCore.dll with no hint that
    it was us. The watchdog may CANCEL the work; it never walks away from it.

    Polls IAsyncInfo rather than registering a completed handler: a handler would mean
    handing WinRT a callback whose lifetime we then have to guarantee across threads,
    and this already runs on a worker thread where blocking is free.
    """
    import time

    info = _qi(op, IID_IAsyncInfo)
    if not info:
        return None, True
    status = ctypes.c_int(ASYNC_STARTED)

    def stillRunning():
        """False once the operation is known to have stopped. Raises if unknowable."""
        if _call(info, SLOT_ASYNCINFO_STATUS, [ctypes.POINTER(ctypes.c_int)],
                 ctypes.byref(status)) != S_OK:
            raise OSError("could not read async status")
        return status.value == ASYNC_STARTED

    try:
        try:
            deadline = time.time() + watchdog
            while stillRunning() and time.time() < deadline:
                time.sleep(0.002)
            if status.value == ASYNC_STARTED:
                # Overdue. Cancel it and wait for it to actually stop.
                log.debugWarning("Spatial Typing Feedback: cancelling overdue operation")
                _call(info, SLOT_ASYNCINFO_CANCEL, [])
                grace = time.time() + 5.0
                while stillRunning() and time.time() < grace:
                    time.sleep(0.002)
        except OSError:
            return None, False
        if status.value == ASYNC_STARTED:
            # It refused to stop. Leaking a reference costs a few kilobytes; freeing
            # memory the engine is still writing into costs the user their screen reader.
            log.error(
                "Spatial Typing Feedback: operation would not cancel - leaking it rather "
                "than freeing memory that is still in use",
            )
            return None, False
        if status.value != ASYNC_COMPLETED:
            log.debugWarning("Spatial Typing Feedback: async status %d" % status.value)
            return None, True
        out = ctypes.c_void_p()
        if _call(op, resultsSlot, [ctypes.POINTER(ctypes.c_void_p)], ctypes.byref(out)) != S_OK:
            return None, True
        return (out if out.value else None), True
    finally:
        _release(info)


class WinRTVoice(object):
    """A Windows SpeechSynthesizer of our own, rendered into panned players.

    Everything COM lives on one worker thread which owns the synthesizer outright. That
    removes any question of apartments and cross-thread marshalling, and synthesis must
    not block NVDA's main thread on every keystroke anyway.
    """

    def __init__(self, name="secondary"):
        self.name = name
        self.available = False
        self.spokeSuccessfully = False
        self.lastError = None
        self.supportsProsodyOptions = True
        self.supportsPunctuationSilence = True
        self.player = PannedPlayer(name)
        #: (id, displayName) for each installed voice. Read on the worker at start-up so
        #: the settings ring can ask for it from the main thread without touching COM.
        self.voices = []
        self._voiceLanguages = {}
        self._queue = None
        self._thread = None
        self._stopping = False
        self._ready = threading.Event()
        self._synth = None
        self._options2 = None
        self._options3 = None
        self._bufferFactory = None
        #: Prosody, expressed per utterance in SSML.
        self._prosodyRate = 100
        self._pitchPercent = 0
        self._volumePercent = 100
        #: Language for the SSML envelope; replaced with the chosen voice's own.
        self._lang = "en-US"

    # -- lifecycle --------------------------------------------------------------

    def initialize(self):
        self._queue = queue.Queue()
        self._stopping = False
        self._ready.clear()
        self._thread = threading.Thread(
            target=self._worker, name="spatialTypingFeedback-%s" % self.name, daemon=True,
        )
        self._thread.start()
        # Wait for the worker to report whether COM setup worked, so initialize() can
        # answer truthfully rather than optimistically.
        self._ready.wait(timeout=10.0)
        return self.available

    def terminate(self):
        self._stopping = True
        self.available = False
        if self._queue is not None:
            try:
                self._queue.put_nowait(_STOP)
            except Exception:  # noqa: BLE001
                pass
        if self._thread is not None and self._thread.is_alive():
            self._thread.join(timeout=3.0)
        self._thread = None
        self._queue = None
        self.player.terminate()

    # -- commands, all executed on the worker -----------------------------------

    def _post(self, *command):
        if self._queue is not None and not self._stopping:
            self._queue.put(command)

    def speak(self, text, player=None, kind=None, interrupt=False):
        if not self.available or not text:
            return
        if interrupt:
            self.cancelKind(kind)
        self._post("speak", text, player if player is not None else self.player, kind)

    def cancelKind(self, kind):
        """Drop pending utterances of one kind, leaving the others alone.

        Characters must interrupt characters or the echo falls behind the keyboard, but
        a character must never cancel a word queued behind it.
        """
        if self._queue is None:
            return
        keep = []
        try:
            while True:
                item = self._queue.get_nowait()
                if item is _STOP or item[0] != "speak" or item[3] != kind:
                    keep.append(item)
                else:
                    item[2].stop()
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
                if item is not _STOP and item[0] == "speak":
                    item[2].stop()
        except queue.Empty:
            pass
        self.player.stop()

    def setRatePercent(self, percent, rateBoost=True):
        """Speed from 0-100. Rate boost widens the range, exactly as NVDA's does.

        Applied through SSML rather than SpeechSynthesizerOptions.SpeakingRate. That
        property stores and reads back correctly but has no effect on the synthesised
        audio - measured, not assumed: identical byte counts at 0.5x, 1.0x and 3.5x.
        SSML prosody does work, and scales exactly as expected.
        """
        percent = max(0, min(100, int(percent)))
        if rateBoost:
            low, high = PROSODY_RATE_MIN, PROSODY_RATE_MAX
        else:
            low, high = PROSODY_RATE_PLAIN_MIN, PROSODY_RATE_PLAIN_MAX
        self._prosodyRate = int(round(low + (percent / 100.0) * (high - low)))
        log.debug(
            "Spatial Typing Feedback: %s speed %d%% (boost=%s) -> prosody rate %d%%"
            % (self.name, percent, rateBoost, self._prosodyRate),
        )
        return True

    def setPitchPercent(self, percent):
        """Pitch from 0-100, 50 being the voice's own. Also via SSML, for the same reason."""
        percent = max(0, min(100, int(percent)))
        # SSML pitch is relative, so map the middle of the range onto no change.
        self._pitchPercent = int(round((percent - 50) * 1.0))
        return True

    def setVolume(self, volume):
        # The stream's channel gain is the real level control; this stays for interface
        # parity and is applied in SSML alongside the rest.
        self._volumePercent = max(0, min(100, int(volume)))
        return True

    def setRate(self, rate, rateBoost=True):
        return self.setRatePercent(rate, rateBoost)

    def setPitch(self, pitch):
        return self.setPitchPercent(pitch)

    def setPunctuationSilence(self, enable):
        # 0 = Default (keep the pauses), 1 = Min.
        self._post("punctuation", 0 if enable else 1)
        return True

    def getAvailableVoiceIds(self):
        return [(voiceId, index) for index, (voiceId, _name) in enumerate(self.voices)]

    def getVoiceDisplayName(self, voiceId):
        for vid, name in self.voices:
            if vid == voiceId:
                return name
        return None

    def matchVoice(self, voiceId):
        """Select a voice by ID.

        The IDs are the same registry paths NVDA reports for its OneCore voices, because
        it is the same voice list, so matching the main synth is a direct comparison.
        """
        if not voiceId:
            return False
        if not any(vid == voiceId for vid, _name in self.voices):
            log.debugWarning("Spatial Typing Feedback: voice %r not installed" % voiceId)
            return False
        self._post("voice", voiceId)
        return True

    # -- the worker -------------------------------------------------------------

    def _worker(self):
        try:
            combase.RoInitialize(RO_INIT_MULTITHREADED)
        except Exception:  # noqa: BLE001
            pass
        try:
            self._setUp()
            self.available = True
            log.info(
                "Spatial Typing Feedback: %s using Windows SpeechSynthesizer directly "
                "(%d voices)" % (self.name, len(self.voices)),
            )
        except Exception as e:  # noqa: BLE001
            self.lastError = str(e)
            log.error(
                "Spatial Typing Feedback: could not start the Windows speech synthesizer",
                exc_info=True,
            )
        finally:
            self._ready.set()

        while not self._stopping:
            try:
                command = self._queue.get()
            except Exception:  # noqa: BLE001
                break
            if command is _STOP or self._stopping:
                break
            try:
                self._handle(command)
            except Exception:  # noqa: BLE001
                log.error("Spatial Typing Feedback: error in %r" % (command[0],), exc_info=True)
        self._tearDown()

    def _setUp(self):
        with _HString(CLASS_SYNTH) as cls:
            synth = ctypes.c_void_p()
            hr = combase.RoActivateInstance(cls, ctypes.byref(synth))
            if hr != S_OK or not synth.value:
                raise OSError("RoActivateInstance failed: 0x%08X" % (hr & 0xFFFFFFFF))
            self._synth = synth

            synth2 = _qi(synth, IID_ISpeechSynthesizer2)
            if synth2:
                options = ctypes.c_void_p()
                if _call(synth2, SLOT_SYNTH2_GET_OPTIONS, [ctypes.POINTER(ctypes.c_void_p)],
                         ctypes.byref(options)) == S_OK and options.value:
                    self._options2 = _qi(options, IID_ISpeechSynthesizerOptions2)
                    self._options3 = _qi(options, IID_ISpeechSynthesizerOptions3)
                    _release(options)
                _release(synth2)
            self.supportsProsodyOptions = self._options2 is not None
            self.supportsPunctuationSilence = self._options3 is not None

            statics = ctypes.c_void_p()
            if combase.RoGetActivationFactory(
                cls, ctypes.byref(_guid(IID_ISpeechSynthesizerStatics)), ctypes.byref(statics),
            ) == S_OK and statics.value:
                self.voices = self._readVoices(statics)
                _release(statics)

        with _HString(CLASS_BUFFER) as bufCls:
            factory = ctypes.c_void_p()
            if combase.RoGetActivationFactory(
                bufCls, ctypes.byref(_guid(IID_IBufferFactory)), ctypes.byref(factory),
            ) == S_OK and factory.value:
                self._bufferFactory = factory

    def _readVoices(self, statics):
        vec = ctypes.c_void_p()
        if _call(statics, SLOT_STATICS_ALL_VOICES, [ctypes.POINTER(ctypes.c_void_p)],
                 ctypes.byref(vec)) != S_OK or not vec.value:
            return []
        result = []
        try:
            count = ctypes.c_uint32()
            _call(vec, SLOT_VECTOR_GET_SIZE, [ctypes.POINTER(ctypes.c_uint32)],
                  ctypes.byref(count))
            for i in range(count.value):
                item = ctypes.c_void_p()
                if _call(vec, SLOT_VECTOR_GET_AT,
                         [ctypes.c_uint32, ctypes.POINTER(ctypes.c_void_p)],
                         ctypes.c_uint32(i), ctypes.byref(item)) != S_OK:
                    continue
                info = _qi(item, IID_IVoiceInformation)
                if info:
                    name = _getString(info, SLOT_VOICE_DISPLAY_NAME)
                    voiceId = _getString(info, SLOT_VOICE_ID)
                    if voiceId:
                        lang = _getString(info, SLOT_VOICE_LANGUAGE)
                        result.append((voiceId, name or voiceId))
                        self._voiceLanguages[voiceId] = lang or "en-US"
                    _release(info)
                _release(item)
        finally:
            _release(vec)
        return result

    def _tearDown(self):
        for ptr in (self._options2, self._options3, self._bufferFactory, self._synth):
            _release(ptr)
        self._options2 = self._options3 = self._bufferFactory = self._synth = None
        try:
            combase.RoUninitialize()
        except Exception:  # noqa: BLE001
            pass

    def _handle(self, command):
        action = command[0]
        if action == "speak":
            self._doSpeak(command[1], command[2])
        elif action == "punctuation" and self._options3:
            _call(self._options3, SLOT_OPT3_PUT_PUNCTUATION, [ctypes.c_int],
                  ctypes.c_int(command[1]))
        elif action == "voice":
            self._doSetVoice(command[1])

    def _doSetVoice(self, voiceId):
        with _HString(CLASS_SYNTH) as cls:
            statics = ctypes.c_void_p()
            if combase.RoGetActivationFactory(
                cls, ctypes.byref(_guid(IID_ISpeechSynthesizerStatics)), ctypes.byref(statics),
            ) != S_OK or not statics.value:
                return
        try:
            vec = ctypes.c_void_p()
            if _call(statics, SLOT_STATICS_ALL_VOICES, [ctypes.POINTER(ctypes.c_void_p)],
                     ctypes.byref(vec)) != S_OK or not vec.value:
                return
            # The collection holds a reference to every installed voice, so returning
            # from the middle of the loop without releasing it leaks all of them - once
            # per voice change, which is a thing the user does repeatedly from the ring.
            try:
                count = ctypes.c_uint32()
                _call(vec, SLOT_VECTOR_GET_SIZE, [ctypes.POINTER(ctypes.c_uint32)],
                      ctypes.byref(count))
                for i in range(count.value):
                    item = ctypes.c_void_p()
                    if _call(vec, SLOT_VECTOR_GET_AT,
                             [ctypes.c_uint32, ctypes.POINTER(ctypes.c_void_p)],
                             ctypes.c_uint32(i), ctypes.byref(item)) != S_OK:
                        continue
                    try:
                        info = _qi(item, IID_IVoiceInformation)
                        if not info:
                            continue
                        try:
                            found = _getString(info, SLOT_VOICE_ID) == voiceId
                        finally:
                            _release(info)
                        if found:
                            _call(self._synth, SLOT_SYNTH_PUT_VOICE, [ctypes.c_void_p], item)
                            self._lang = self._voiceLanguages.get(voiceId, "en-US")
                            log.info("Spatial Typing Feedback: %s voice set" % self.name)
                            return
                    finally:
                        _release(item)
            finally:
                _release(vec)
        finally:
            _release(statics)

    def _buildSsml(self, text):
        """Wrap text with the prosody the user has chosen.

        Everything about how the voice sounds is expressed here, because the Options
        properties turned out not to affect synthesis.
        """
        escaped = (
            text.replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;")
        )
        attrs = ['rate="%d%%"' % self._prosodyRate]
        if self._pitchPercent:
            attrs.append('pitch="%+d%%"' % self._pitchPercent)
        if self._volumePercent != 100:
            attrs.append('volume="%d%%"' % self._volumePercent)
        return (
            '<speak version="1.0" xmlns="http://www.w3.org/2001/10/synthesis" '
            'xml:lang="%s"><prosody %s>%s</prosody></speak>'
            % (self._lang, " ".join(attrs), escaped)
        )

    def _doSpeak(self, text, player):
        # The SSML string must outlive the asynchronous call, not the statement that
        # starts it. Freeing it early corrupts the speech engine's heap and takes the
        # whole process down.
        ssml = _HString(self._buildSsml(text))
        stopped = True
        try:
            op = ctypes.c_void_p()
            hr = _call(self._synth, SLOT_SYNTH_SSML_ASYNC,
                       [HSTRING, ctypes.POINTER(ctypes.c_void_p)], ssml.handle,
                       ctypes.byref(op))
            if hr != S_OK or not op.value:
                log.error(
                    "Spatial Typing Feedback: synthesis failed 0x%08X" % (hr & 0xFFFFFFFF),
                )
                return
            stream, stopped = _await(op, SLOT_OP_GET_RESULTS)
            if stopped:
                _release(op)
        finally:
            # Only safe once synthesis has stopped reading it.
            if stopped:
                ssml.free()
        if not stream:
            return
        try:
            data = self._readStream(stream)
        finally:
            _release(stream)
        if not data:
            return
        self.spokeSuccessfully = True
        # The stream is a complete RIFF WAVE, so let the wave module find the PCM rather
        # than assuming a header length.
        try:
            with wave.open(io.BytesIO(data), "r") as wv:
                rate, channels = wv.getframerate(), wv.getnchannels()
                frames = wv.readframes(wv.getnframes())
        except Exception:  # noqa: BLE001
            log.error("Spatial Typing Feedback: could not parse synthesised audio", exc_info=True)
            return
        if channels == 1:
            player.feedMono(frames, rate)
        else:
            player.feedStereo(frames, rate)

    def _readStream(self, stream):
        ras = _qi(stream, IID_IRandomAccessStream)
        if not ras:
            return None
        try:
            size = ctypes.c_uint64()
            if _call(ras, SLOT_STREAM_GET_SIZE, [ctypes.POINTER(ctypes.c_uint64)],
                     ctypes.byref(size)) != S_OK or not size.value:
                return None
            if not self._bufferFactory:
                return None
            buf = ctypes.c_void_p()
            if _call(self._bufferFactory, SLOT_BUFFERFACTORY_CREATE,
                     [ctypes.c_uint32, ctypes.POINTER(ctypes.c_void_p)],
                     ctypes.c_uint32(size.value), ctypes.byref(buf)) != S_OK:
                return None
            bufSafe = True
            inp = _qi(stream, IID_IInputStream)
            if not inp:
                _release(buf)
                return None
            try:
                readOp = ctypes.c_void_p()
                if _call(inp, SLOT_INPUTSTREAM_READ_ASYNC,
                         [ctypes.c_void_p, ctypes.c_uint32, ctypes.c_int,
                          ctypes.POINTER(ctypes.c_void_p)],
                         buf, ctypes.c_uint32(size.value), ctypes.c_int(0),
                         ctypes.byref(readOp)) != S_OK:
                    return None
                filled, stopped = _await(readOp, SLOT_OP_PROGRESS_GET_RESULTS)
                if stopped:
                    _release(readOp)
                else:
                    bufSafe = False
                    return None
                if not filled:
                    return None
                try:
                    length = ctypes.c_uint32()
                    _call(filled, SLOT_BUFFER_GET_LENGTH, [ctypes.POINTER(ctypes.c_uint32)],
                          ctypes.byref(length))
                    access = _qi(filled, IID_IBufferByteAccess)
                    if not access:
                        return None
                    try:
                        raw = ctypes.POINTER(ctypes.c_ubyte)()
                        _call(access, SLOT_BYTEACCESS_BUFFER,
                              [ctypes.POINTER(ctypes.POINTER(ctypes.c_ubyte))], ctypes.byref(raw))
                        return ctypes.string_at(raw, length.value)
                    finally:
                        _release(access)
                finally:
                    _release(filled)
            finally:
                _release(inp)
                if bufSafe:
                    _release(buf)
        finally:
            _release(ras)
