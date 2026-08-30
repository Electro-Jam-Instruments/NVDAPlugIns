# Research: NVDA audio panning, typing echo and error hooks

Verified against `github.com/nvaccess/nvda` **master**, fetched 2026-08-30.
Line numbers are from that snapshot and will drift - treat them as a starting point,
not a contract.

---

## 1. Panning: `nvwave.WavePlayer`

`source/nvwave.py`

```python
class WavePlayer:
    def __init__(
        self,
        channels: int,
        samplesPerSec: int,
        bitsPerSample: int,
        outputDevice: str = DEFAULT_DEVICE_KEY,
        wantDucking: bool = True,
        purpose: AudioPurpose = AudioPurpose.SPEECH,
    ): ...
```

Public methods: `open()`, `close()`, `feed(data, size=None, onDone=None)`, `sync()`,
`idle()`, `stop()`, `pause(switch)`, `setVolume(...)`,
`enableTrimmingLeadingSilence(enable)`, `startTrimmingLeadingSilence(start=True)`.

### The panning primitive (nvwave.py:416)

```python
def setVolume(
    self,
    *,
    all: float | None = None,
    left: float | None = None,
    right: float | None = None,
):
    """Set the volume of one or more channels in this stream.
    Levels must be specified as a number between 0 and 1."""
    if all is None and left is None and right is None:
        raise ValueError("At least one of all, left or right must be specified")
    if all is not None:
        if left is not None or right is not None:
            raise ValueError("all specified, so left and right must not be specified")
        left = right = all
    wasapi.wasPlay_setChannelVolume(self._player, 0, c_float(left))
    try:
        wasapi.wasPlay_setChannelVolume(self._player, 1, c_float(right))
    except OSError as e:
        # E_INVALIDARG indicates that the audio device doesn't support this channel.
        if not (all and e.winerror == E_INVALIDARG):
            raise
```

Notes:
- Keyword-only arguments.
- **`open()` and `stop()` can wipe panning.** Both call `_setVolumeFromConfig`, which
  calls `setVolume(all=...)` - equal on both channels:

  ```python
  def _setVolumeFromConfig(self):
      if self._purpose is not AudioPurpose.SOUNDS:
          return
      volume = config.conf["audio"]["soundVolume"]
      if config.conf["audio"]["soundVolumeFollowsVoice"]:
          synth = synthDriverHandler.getSynth()
          if synth and synth.isSupported("volume"):
              volume = synth.volume
      self.setVolume(all=volume / 100)
  ```

  The early return means **speech** players (`AudioPurpose.SPEECH`, the default) are
  safe. **Sound** players (`AudioPurpose.SOUNDS`) are not - our error alert must
  re-apply its pan before each `feed()`, folding `soundVolume` into the coefficients.
- `WavePlayer._instances` is a `weakref.WeakValueDictionary` mapping native player
  pointers to Python instances. A registry exists, but identifying *which* entry belongs
  to the main synth from it is guesswork - go via `getSynth()._player` instead (D9).
- A mono device raises `E_INVALIDARG` on channel 1. Panning is meaningless there;
  detect and degrade gracefully rather than crashing.
- **NVDA 2026.1 renamed `WasapiWavePlayer` to `WavePlayer`** and changed `__init__` so
  `outputDevice` takes only string arguments. Any version guard must account for both
  names.

### Module-level

```python
def playWaveFile(fileName: str, asynchronous: bool = True,
                 isSpeechWaveFileCommand: bool = False)
```

Crucially it consults an extension point before playing (nvwave.py:105):

```python
if not decide_playWaveFile.decide(
    fileName=fileName,
    asynchronous=asynchronous,
    isSpeechWaveFileCommand=isSpeechWaveFileCommand,
):
    log.debug("Playing wave file canceled by handler registered to "
              "decide_playWaveFile extension point")
    return
```

**This is our supported hook for R2.** Register a
handler, match on `fileName`, return `False` to cancel NVDA's playback, and play the
file ourselves through a panned `WavePlayer`.

Also note `playWaveFile` maintains a single global `fileWavePlayer` and stops any
in-progress wave before starting a new one. Our own player is separate, so our panned
sounds will not be cut off by NVDA's next wave - and vice versa. That is a behaviour
change worth watching for during testing.

---

## 2. Keyboard echo (R1)

### Call chain

`source/NVDAObjects/__init__.py`

```python
def event_typedCharacter(self, ch):
    speech.speakTypedCharacters(ch)
    import winUser
    if (
        config.conf["keyboard"]["beepForLowercaseWithCapslock"]
        and ch.islower()
        and winUser.getKeyState(winUser.VK_CAPITAL) & 1
    ):
        import tones
        tones.beep(3000, 40)
```

### `speech/speech.py:1437`

```python
def speakTypedCharacters(ch: str):
    typingIsProtected = api.isTypingProtected()
    if typingIsProtected:
        realChar = PROTECTED_CHAR
    else:
        realChar = ch
    if unicodedata.category(ch)[0] in "LMN":
        _curWordChars.append(realChar)
    elif ch == "\b":
        del _curWordChars[-1:]          # backspace
    elif ch == "":
        return                          # control+backspace in some apps
    elif len(_curWordChars) > 0:
        typedWord = "".join(_curWordChars)
        clearTypedWordBuffer()
        typingEchoMode = config.conf["keyboard"]["speakTypedWords"]
        if typingEchoMode != TypingEcho.OFF.value and not typingIsProtected:
            if typingEchoMode == TypingEcho.ALWAYS.value or (
                typingEchoMode == TypingEcho.EDIT_CONTROLS.value and isFocusEditable()
            ):
                speakText(typedWord)
    if _speechState._suppressSpeakTypedCharactersNumber > 0:
        suppress = time.time() - _speechState._suppressSpeakTypedCharactersTime <= 0.1
        if suppress:
            _speechState._suppressSpeakTypedCharactersNumber -= 1
        else:
            _speechState._suppressSpeakTypedCharactersNumber = 0
            _speechState._suppressSpeakTypedCharactersTime = None
    else:
        suppress = False

    typingEchoMode = config.conf["keyboard"]["speakTypedCharacters"]
    if not suppress and typingEchoMode != TypingEcho.OFF.value and ch >= FIRST_NONCONTROL_CHAR:
        if typingEchoMode == TypingEcho.ALWAYS.value or (
            typingEchoMode == TypingEcho.EDIT_CONTROLS.value and isFocusEditable()
        ):
            speakSpelling(realChar)
```

Exported from `speech/__init__.py` (line 63 in the `from .speech import (...)` block,
and listed in `__all__` at line 144), so `speech.speakTypedCharacters` is the
attribute every caller resolves.

### Config values

`source/config/configSpec.py`, `[keyboard]`:

```
speakTypedCharacters = integer(default=1,min=0,max=2)
speakTypedWords      = integer(default=0,min=0,max=2)
beepForLowercaseWithCapslock = boolean(default=true)
alertForSpellingErrors = boolean(default=True)
```

`source/config/configFlags.py:55`:

```python
class TypingEcho(DisplayStringIntEnum):
    OFF = 0
    EDIT_CONTROLS = 1
    ALWAYS = 2
```

### Terminal special case

`source/NVDAObjects/behaviors.py:622` - terminals buffer typed characters rather than
echoing them immediately, and call `speech.clearTypedWordBuffer()` on tab. Any
reimplementation of the echo logic must not double-speak here.

---

## 3. Typing error alert (R2)

`source/NVDAObjects/behaviors.py:306` (`EditableTextBase.event_typedCharacter`):

```python
def event_typedCharacter(self, ch: str):
    if (
        config.conf["documentFormatting"]["reportSpellingErrors2"] != ReportSpellingErrors.OFF.value
        and config.conf["keyboard"]["alertForSpellingErrors"]
        and (
            # Not alpha, apostrophe or control.
            ch.isspace() or (ch >= " " and ch not in "'\x7f" and not ch.isalpha())
        )
    ):
        self._reportErrorInPreviousWord()
    super().event_typedCharacter(ch)
```

`_reportErrorInPreviousWord` (behaviors.py:263) takes a caret TextInfo, moves back two
characters, then defers the actual check by 50ms because MS Word's UIA needs time to
mark the error:

```python
def _delayedDetection():
    fields = info.getTextWithFields()
    for command in fields:
        if (isinstance(command, textInfos.FieldCommand)
                and command.command == "formatChange"
                and command.field.get("invalid-spelling")):
            break
    else:
        return  # No error.
    if speech.getState().speechMode not in [speech.SpeechMode.off, speech.SpeechMode.onDemand]:
        nvwave.playWaveFile(os.path.join(globalVars.appDir, "waves", "textError.wav"))

core.callLater(50, _delayedDetection)
```

**Hook:** `decide_playWaveFile` matching `textError.wav` **with
`isSpeechWaveFileCommand=False`**. All detection logic stays in NVDA.

### `reportSpellingErrors2` is a bitmask, not a toggle

`source/config/configFlags.py:185`:

```python
class ReportSpellingErrors(DisplayStringIntFlag):
    OFF = 0b0
    SPEECH = 0b1
    SOUND = 0b10
    SPEECH_AND_SOUND = SPEECH | SOUND
    BRAILLE = 0b100
```

Hence `integer(min=0, max=7, default=1)` in configSpec - `SPEECH|SOUND|BRAILLE`.
The typing-time check above only tests `!= OFF`, so it fires whenever *any* reporting
mode is on.

### NVDA does speak errors - but only while reading, not while typing

`source/speech/speech.py:3059` (`getFormatFieldSpeech`):

```python
if formatConfig["reportSpellingErrors2"]:
    invalidSpelling = attrs.get("invalid-spelling")
    oldInvalidSpelling = attrsCache.get("invalid-spelling") if attrsCache is not None else None
    if (invalidSpelling or oldInvalidSpelling is not None) and invalidSpelling != oldInvalidSpelling:
        texts = []
        if invalidSpelling:
            if formatConfig["reportSpellingErrors2"] & ReportSpellingErrors.SOUND.value:
                texts.append(WaveFileCommand(r"waves	extError.wav"))
            if formatConfig["reportSpellingErrors2"] & ReportSpellingErrors.SPEECH.value:
                # Translators: Reported when text contains a spelling error.
                texts.append(_("spelling error"))
        elif extraDetail and _shouldReportOutOfError(formatConfig):
            # Translators: Reported when moving out of text containing a spelling error.
            texts.append(_("out of spelling error"))
        textList.extend(texts)
```

The same block handles `invalid-grammar` -> `_("grammar error")` /
`_("out of grammar error")`.

So there are **three** error presentations, and they are distinguishable:

| # | Trigger | Mechanism | `isSpeechWaveFileCommand` |
|---|---------|-----------|---------------------------|
| 1 | Typing a word-ending character | `nvwave.playWaveFile(textError.wav)` | `False` |
| 2 | Reading over marked text, SOUND bit | `WaveFileCommand` in the speech sequence | `True` |
| 3 | Reading over marked text, SPEECH bit | literal text in the speech sequence | n/a |

`source/speech/commands.py:458`:

```python
class WaveFileCommand(BaseCallbackCommand):
    """Play a wave file."""
    def __init__(self, fileName):
        self.fileName = fileName

    def run(self):
        import nvwave
        nvwave.playWaveFile(self.fileName, asynchronous=True, isSpeechWaveFileCommand=True)
```

Only #1 is a typing event. #2 and #3 annotate text the main voice is reading and must
stay in sequence with it (Decision D5).

### Bonus: `tones.beep` is already pannable

`source/speech/commands.py`, `BeepCommand.run`:

```python
tones.beep(self.hz, self.length, left=self.left, right=self.right, isSpeechBeepCommand=True)
```

If the caps-lock beep or indentation tones ever need positioning, `tones.beep` takes
`left=` and `right=` directly - no wave player needed.

---

## 4. Word echo - completed words (R3)

R3 is NVDA's **word echo**, and it lives in the same function as the character echo -
see section 2 above. The relevant branch of `speakTypedCharacters` is:

```python
elif len(_curWordChars) > 0:
    typedWord = "".join(_curWordChars)
    clearTypedWordBuffer()
    typingEchoMode = config.conf["keyboard"]["speakTypedWords"]
    if typingEchoMode != TypingEcho.OFF.value and not typingIsProtected:
        if typingEchoMode == TypingEcho.ALWAYS.value or (
            typingEchoMode == TypingEcho.EDIT_CONTROLS.value and isFocusEditable()
        ):
            speakText(typedWord)
```

A word "completes" when the typed character is not a letter, mark or number
(`unicodedata.category(ch)[0] in "LMN"` fails) and the buffer is non-empty.

`FIRST_NONCONTROL_CHAR = " "` (speech.py:1423), so a space passes the character-echo
gate too. One space therefore fires **both** branches: the completed word via
`speakText`, and the space itself via `speakSpelling`. No extra hook is needed for R3 -
it is the same wrap as R1, routed to a different pan position.

---

## 4b. Application autocomplete - OUT OF SCOPE, retained for reference

Not required by R3. Recorded because an earlier draft confused the two, and because it
may become a separate feature later.

`source/NVDAObjects/behaviors.py:208`:

```python
class InputFieldWithSuggestions(NVDAObject):
    """Allows NVDA to announce appearance/disappearance of suggestions as content
    is entered. This is used in various places, including Windows 10 search edit
    fields and others."""

    def event_suggestionsOpened(self):
        braille.handler.message(_("Suggestions"))
        if config.conf["presentation"]["reportAutoSuggestionsWithSound"]:
            nvwave.playWaveFile(os.path.join(globalVars.appDir, "waves", "suggestionsOpened.wav"))

    def event_suggestionsClosed(self):
        if config.conf["presentation"]["reportAutoSuggestionsWithSound"]:
            nvwave.playWaveFile(os.path.join(globalVars.appDir, "waves", "suggestionsClosed.wav"))

    def event_controllerForChange(self):
        # Report when suggestions appear and disappear.
        if self is api.getFocusObject() and len(self.controllerFor) > 0:
            self.event_suggestionsOpened()
        else:
            self.event_suggestionsClosed()
```

`EditableText` inherits this via `EditableTextWithSuggestions`.

Config: `[presentation] reportAutoSuggestionsWithSound = boolean(default=True)`.

**Sounds:** covered by the same `decide_playWaveFile` hook -
`suggestionsOpened.wav` / `suggestionsClosed.wav`.

**Text:** no dedicated function - this is why autocomplete was ruled out. NVDA 2024.x release notes state that
`SearchField` / `SuggestionListItem` UIA objects are no longer required, because
"automatic reporting of search suggestions ... has been exposed via UI Automation with
the controllerFor pattern ... available generically via behaviours.EditableText and
the base NVDAObject". So `controllerFor` is the modern, generic path - and the right
place to look during Spike S2.

---

## 4c. OneCore as a second, independent instance (D3)

`source/synthDrivers/oneCore.py`

### Token-based, therefore instanceable

```python
ocSpeech_Callback = ctypes.CFUNCTYPE(None, ctypes.c_void_p, ctypes.c_int, ctypes.c_wchar_p)

def __init__(self):
    self._dll = NVDAHelper.getHelperLocalWin10Dll()
    self._dll.ocSpeech_initialize.restype = HANDLE
    ...
    self._callbackInst = ocSpeech_Callback(self._callback)
    self._ocSpeechToken = HANDLE()
    self._ocSpeechToken.value = self._dll.ocSpeech_initialize(self._callbackInst)
    self._dll.ocSpeech_getVoices.restype = comtypes.BSTR
    self._dll.ocSpeech_getCurrentVoiceId.restype = ctypes.c_wchar_p
    self._player = None
```

Every subsequent call passes the token:
`ocSpeech_speak(token, text)`, `ocSpeech_setRate(token, raw)`,
`ocSpeech_setPitch(token, raw)`, `ocSpeech_setVolume(token, raw)`,
`ocSpeech_setVoice(token, index)`, `ocSpeech_getVoices(token)`,
`ocSpeech_getCurrentVoiceId(token)`, `ocSpeech_terminate(token)`.

The capability probes are token-free and can be called before initializing:
`ocSpeech_supportsProsodyOptions() -> bool`,
`ocSpeech_supportsPunctuationSilence() -> bool`.

**Note:** eSpeak's driver *is* module-global. OneCore is not. Do not generalise from one
to the other.

### Audio arrives as WAV bytes on a background thread

```python
WAVE_HEADER_LENGTH = 46

def _callback(self, bytes, len, markers):
    if self._earlyExitCB:
        return
    if len == 0:
        self._handleSpeechFailure()
        return
    # This gets called in a background thread.
    stream = io.BytesIO(ctypes.string_at(bytes, WAVE_HEADER_LENGTH))
    wav = wave.open(stream, "r")
    self._maybeInitPlayer(wav)
    data = bytes + WAVE_HEADER_LENGTH
    dataLen = wav.getnframes() * wav.getnchannels() * wav.getsampwidth()
    ...
    self._player.feed(ctypes.c_void_p(data + prevPos), size=dataLen - prevPos)
```

`_earlyExitCB` is set before `ocSpeech_terminate` so pending callbacks stop touching
the instance. Copy that pattern.

### The player, and why we cannot pan theirs

```python
def _maybeInitPlayer(self, wav):
    samplesPerSec = wav.getframerate()
    if self._player and self._player.samplesPerSec == samplesPerSec:
        return
    if self._player:
        self._player.idle()
    bytesPerSample = wav.getsampwidth()
    self._bytesPerSec = samplesPerSec * bytesPerSample
    self._player = nvwave.WavePlayer(
        channels=wav.getnchannels(),        # mono
        samplesPerSec=samplesPerSec,
        bitsPerSample=bytesPerSample * 8,
        outputDevice=config.conf["audio"]["outputDevice"],
    )
```

`channels` comes from the wave header, and OneCore is mono. Panning a mono player is a
contradiction - `setVolume` raises `E_INVALIDARG` on channel 1 (section 1). We build
our own `channels=2` player and upmix before feeding.

### Rate boost

```python
MIN_PITCH = 0.0
MAX_PITCH = 2.0
MIN_RATE = 0.5
DEFAULT_MAX_RATE = 1.5
BOOSTED_MAX_RATE = 6.0

@classmethod
def _percentToParam(cls, percent, min, max):
    """Overrides SynthDriver._percentToParam to return floating point parameter values."""
    return float(percent) / 100 * (max - min) + min

def _set_rate(self, rate):
    self._rate = rate
    if not self.supportsProsodyOptions:
        return
    maxRate = self.BOOSTED_MAX_RATE if self._rateBoost else self.DEFAULT_MAX_RATE
    rawRate = self._percentToParam(rate, self.MIN_RATE, maxRate)
    self._queuedSpeech.append((self._dll.ocSpeech_setRate, rawRate))
```

Rate boost is entirely this: swap `1.5` for `6.0` as the top of the mapping range. It
is not a separate OneCore mode. `_set_rateBoost` just re-applies the cached rate under
the new range.

All of rate, pitch and volume are gated on `supportsProsodyOptions` (OneCore API > 5).
When false the driver logs `"Prosody options not supported"` and the values are fixed at
initialization - no rate boost for anyone, including NVDA's main synth.

### Voice enumeration

`ocSpeech_getVoices(token)` returns a `BSTR` of `|`-separated voice strings, each
carrying ID, language and display name
(`_getVoiceInfoFromOnecoreVoiceString`). Same list the user sees in NVDA's synthesizer
settings.

---

## 5. Speech extension points

`source/speech/extensions.py`

| Extension point | Type | Handler arguments |
|---|---|---|
| `speechCanceled` | Action | (none) |
| `pre_speechCanceled` | Action | (none) |
| `post_speechPaused` | Action | `switch` (bool) |
| `pre_speech` | Action | `speechSequence`, `symbolLevel`, `priority` |
| `filter_speechSequence` | Filter | `value` (SpeechSequence) |
| `pre_speechQueued` | Action | `speechSequence`, `priority` |

`filter_speechSequence` is the only one that can *change or suppress* text on its way
to the synth. `pre_speech` and `pre_speechQueued` are notifications only.

None of them carry provenance - they cannot tell you *why* something is being spoken.
That limitation is what forces the flag-plus-inference design in Decision D6.

---

## 5b. The synth settings ring (R7 / D8)

`source/synthSettingsRing.py`, `source/globalCommands.py`, `source/synthDriverHandler.py`

### It is one swappable global

```python
# globalCommands.py - every ring script funnels through globalVars.settingsRing
def script_nextSynthSetting(self, gesture):
    nextSettingName = globalVars.settingsRing.next()
    ...
    nextSettingValue = globalVars.settingsRing.currentSettingValue

def script_increaseSynthSetting(self, gesture):
    settingName = globalVars.settingsRing.currentSettingName
    ...
    settingValue = globalVars.settingsRing.increase()
```

Also present: `previous()`, `decrease()`, `increaseLarge()`, `decreaseLarge()`,
`first()`, `last()`, `currentSettingName`, `currentSettingValue`.

### Created once, then updated in place

`synthDriverHandler.py:445`:

```python
# start or update the synthSettingsRing
if globalVars.settingsRing:
    globalVars.settingsRing.updateSupportedSettings(synth)
else:
    globalVars.settingsRing = SynthSettingsRing(synth)
```

**This is the hook.** Replace `globalVars.settingsRing` with our own subclass and NVDA
will keep calling `updateSupportedSettings` on it forever after. Override that method,
call `super()`, then append our slots.

### How the ring builds itself

```python
def updateSupportedSettings(self, synth):
    prevID = (self.settings[self._current].setting.id
              if self._current is not None and hasattr(self, "settings") else None)
    settings: list[SynthSetting] = []
    for s in synth.supportedSettings:
        if not s.availableInSettingsRing:
            continue
        if prevID == s.id:  # restore the last setting
            self._current = len(settings)
        if isinstance(s, NumericDriverSetting):
            cls = SynthSetting
        elif isinstance(s, BooleanDriverSetting):
            cls = BooleanSynthSetting
        else:
            cls = StringSynthSetting
        settings.append(cls(synth, s))
    if len(settings) == 0:
        self._current = None
        self.settings = None
    else:
        self.settings = settings
    ...
```

Gotchas for a subclass:
- `self.settings` is set to **`None`**, not `[]`, when the synth exposes nothing. Handle
  that before appending.
- Ring position is restored by matching `setting.id`, so our slots need stable unique ids.
- Appending after `super()` keeps NVDA's indices unchanged and puts ours at the end.

### Where values are written - and why we must override

```python
class SynthSetting(baseObject.AutoPropertyObject):
    def increase(self):
        val = min(self.max, self.value + self.step)
        self.value = val
        return self._getReportValue(val)

    def _get_value(self):
        return getattr(self.synth, self.setting.id)

    def _set_value(self, value):
        setattr(self.synth, self.setting.id, value)
        config.conf["speech"][self.synth.name][self.setting.id] = value

    def _getReportValue(self, val):
        return str(val)
```

`_set_value` writes into NVDA's per-synth speech config. Our settings do not belong
there and are not in that config spec, so our subclass must override `_set_value` (and
`_get_value`) to use our own config section.

`_getReportValue` is the override that turns a pan number into "left 60" / "centre".

Also note `min`/`max`/`step`/`largeStep` are read from the `NumericDriverSetting` in
`SynthSetting.__init__`, and `increase`/`decrease` are plain clamped arithmetic - a
negative `minVal` such as -100 works without special handling.

### Setting definitions

`source/autoSettingsUtils/driverSetting.py`:

```python
class DriverSetting(AutoPropertyObject):
    def __init__(self, id, displayNameWithAccelerator,
                 availableInSettingsRing: bool = False,
                 defaultVal=None, displayName: str | None = None, useConfig=True): ...

class NumericDriverSetting(DriverSetting):
    def __init__(self, id, displayNameWithAccelerator, availableInSettingsRing=False,
                 defaultVal=50, minVal=0, maxVal=100, minStep=1, normalStep=5,
                 largeStep=10, displayName=None, useConfig=True): ...

class BooleanDriverSetting(DriverSetting): ...
```

`availableInSettingsRing=True` is required or `updateSupportedSettings` skips it.
`useConfig=False` opts out of NVDA's automatic config persistence.

---

## 6. Audio config surface

`source/config/configSpec.py`, `[audio]`:

```
outputDevice = string(default=default)
audioDuckingMode = integer(default=0)
soundVolumeFollowsVoice = boolean(default=false)
soundVolume = integer(default=100, min=0, max=100)
audioAwakeTime = integer(default=30, min=0, max=3600)
whiteNoiseVolume = integer(default=0, min=0, max=100)
soundSplitState = integer(default=0)
includedSoundSplitModes = int_list(default=list(0, 2, 3))
```

`source/audio/__init__.py` re-exports `SoundSplitState`, `_setSoundSplitState()` and
`_toggleSoundSplitState()` from `source/audio/soundSplit.py`. The leading underscores
mark these as private - do not call them.

**Sound Split works on Windows audio sessions, not on wave players:**

```python
class _VolumeSetter(AudioSessionNotification):
    def on_session_created(self, new_session: AudioSession):
        channelVolume = new_session.channelAudioVolume()
        channelVolume.SetChannelVolume(0, self.leftNVDAVolume, None)
        channelVolume.SetChannelVolume(1, self.rightNVDAVolume, None)
```

`_applyToAllAudioSessions` walks every session via `IAudioSessionManager2`. This is one
level above our per-player `setVolume` and applies to the whole NVDA process, so the two
multiply.

`SoundSplitState` is an eight-value `DisplayStringIntEnum` (`OFF`,
`NVDA_LEFT_APPS_RIGHT`, `NVDA_BOTH_APPS_LEFT`, and so on) with a `getNVDAVolume()`
returning the channel pair applied to NVDA's session.

We do not need to reason about individual states. See Decision D7: the two features are
mutually exclusive - if Sound Split is on, this add-on stays dormant.

---

## 7. Open questions for the spikes

**S1 - the blocking question**
- **Can `ocSpeech_initialize` be called a second time in one process while NVDA's own
  OneCore driver holds a live token?** Everything in D3 rests on yes. The token design
  says it should work; it is not proven. Test this first, before any other spike work.
- If two tokens work, do they contend for the Windows speech platform under load
  (fast typing, main voice speaking simultaneously)?
- Latency: time from `ocSpeech_speak` to first PCM in the callback, for a single
  character. Character echo must keep up with fast typing.
- Does a second `WavePlayer` cause WASAPI contention or device-format conflict with the
  main synth's player?
- Confirm the mono-to-stereo upmix is cheap enough at character-echo rates, and that
  `setVolume` then works on the resulting stereo player.
- `wantDucking` on our player: does it interact badly with NVDA's own ducking?
- Mono output *devices*: `setVolume` raises `E_INVALIDARG` on channel 1. Confirm the
  degradation path.

**S1 additions from the R3 clarification**
- One space fires both echo branches at once. Can the secondary voice play a word
  centred and a character panned right simultaneously, or does one have to queue behind
  the other? Two `WavePlayer` instances may be needed rather than one.
- If they must queue, which wins? Speaking the word first is probably right, but this
  needs listening to, not reasoning about.

---

## Sources

- [nvaccess/nvda source (master)](https://github.com/nvaccess/nvda)
- [NVDA Developer Guide](https://download.nvaccess.org/documentation/developerGuide.html)
- [What's New in NVDA 2026.1](https://download.nvaccess.org/releases/2026.1/documentation/changes.html)
- [nvwave.WavePlayer API reference](https://www.webbie.org.uk/nvda/api/nvwave.WavePlayer-class.html)
- [NVDA 2024.2 Developer Guide](https://download.nvaccess.org/releases/2024.2/documentation/developerGuide.html)
