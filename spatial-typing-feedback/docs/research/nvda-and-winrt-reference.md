# Reference: the NVDA and WinRT surfaces this add-on uses

Only what the shipped code depends on. Everything here was verified against a running
system or against `nvaccess/nvda` source, not taken from documentation - in several places
the documentation and the behaviour disagree, and those are called out.

Line numbers drift; treat them as a starting point.

---

## 1. NVDA: audio output

`source/nvwave.py`

```python
def setVolume(self, *, all=None, left=None, right=None):
    """Levels must be specified as a number between 0 and 1."""
    ...
    wasapi.wasPlay_setChannelVolume(self._player, 0, c_float(left))
    try:
        wasapi.wasPlay_setChannelVolume(self._player, 1, c_float(right))
    except OSError as e:
        # E_INVALIDARG indicates that the audio device doesn't support this channel.
```

- Keyword-only. `all` cannot be combined with `left`/`right`.
- **A mono player raises `E_INVALIDARG` on channel 1.** Always create with `channels=2`
  and upmix.
- **`open()` and `stop()` can wipe panning.** Both call `_setVolumeFromConfig`, which
  calls `setVolume(all=...)` - but it returns early unless
  `purpose is AudioPurpose.SOUNDS`. Speech players are therefore safe; sound players are
  not.

### Class name changed

`WasapiWavePlayer` in NVDA 2024.1, `WavePlayer` from 2025.1. Both names are handled.

### Wave file interception

```python
if not decide_playWaveFile.decide(
    fileName=fileName, asynchronous=asynchronous,
    isSpeechWaveFileCommand=isSpeechWaveFileCommand,
):
    return
```

Returning `False` cancels NVDA's playback. `isSpeechWaveFileCommand` is how a typing-time
alert is told from one embedded in a speech sequence.

---

## 2. NVDA: typing echo

`NVDAObject.event_typedCharacter` -> `speech.speakTypedCharacters(ch)`, which decides both
echoes and owns the word buffer:

```python
elif len(_curWordChars) > 0:
    typedWord = "".join(_curWordChars)
    clearTypedWordBuffer()
    if speakTypedWords is on:
        speakText(typedWord)          # word echo
...
if not suppress and speakTypedCharacters is on and ch >= FIRST_NONCONTROL_CHAR:
    speakSpelling(realChar)           # character echo
```

- `FIRST_NONCONTROL_CHAR = " "`, so a space fires **both** branches.
- `TypingEcho` is `OFF=0, EDIT_CONTROLS=1, ALWAYS=2` - the order is easy to get backwards.
- Terminals buffer characters instead of echoing them and call
  `speech.clearTypedWordBuffer()` themselves.

We wrap `speech.speakTypedCharacters` and swap `speech.speech.speakText` /
`speakSpelling` for the duration of the call, so all of that state stays with NVDA.

---

## 3. NVDA: the typing-error alert

`NVDAObjects/behaviors.py`, `EditableTextBase.event_typedCharacter`:

```python
if (config.conf["documentFormatting"]["reportSpellingErrors2"] != OFF
    and config.conf["keyboard"]["alertForSpellingErrors"]
    and (ch.isspace() or (ch >= " " and ch not in "'\x7f" and not ch.isalpha()))):
    self._reportErrorInPreviousWord()
```

`_reportErrorInPreviousWord` waits 50 ms for the application to mark the error, checks for
an `invalid-spelling` format field, then plays `waves/textError.wav`. All of that stays
with NVDA; we only change where the sound comes from.

`reportSpellingErrors2` is a **bitmask**, not a toggle: `OFF=0, SPEECH=1, SOUND=2,
BRAILLE=4`.

---

## 4. NVDA: the settings ring

`globalVars.settingsRing` is one swappable object. Every ring script in `globalCommands`
goes through it, and `synthDriverHandler` updates the existing object rather than
replacing it:

```python
if globalVars.settingsRing:
    globalVars.settingsRing.updateSupportedSettings(synth)
else:
    globalVars.settingsRing = SynthSettingsRing(synth)
```

So a subclass installed once survives synth changes.

Gotchas:

- `self.settings` is set to **`None`**, not `[]`, when the synth exposes nothing.
- `SynthSetting._set_value` writes to `config.conf["speech"][synthName][id]`. Settings
  that are not the synth's must override it.
- Ring position is restored by matching `setting.id`, so ids must be stable and unique.
- `NumericDriverSetting` needs `availableInSettingsRing=True` or the ring skips it.

`speech.extensions.filter_speechSequence` is applied in `speak()`, which returns
immediately on an empty sequence - so returning `[]` suppresses an utterance. **It does
not exist in NVDA 2024.1**, which is why the minimum supported version is 2025.1.

---

## 5. Windows: activating a SpeechSynthesizer directly

Every IID and vtable slot below was read off the running system with
`IInspectable::GetIids` and confirmed with `GetRuntimeClassName`. One of them differs from
the value commonly quoted, which is why none of them were taken on trust.

### Interface IDs

| Interface | IID |
|-----------|-----|
| `ISpeechSynthesizer` (default) | `{CE9F7C76-97F4-4CED-AD68-D51C458E45C6}` |
| `ISpeechSynthesizer2` | `{A7C5ECB2-4339-4D6A-BBF8-C7A4F1544C2E}` |
| `ISpeechSynthesizerStatics` | `{7D526ECC-7533-4C3F-85BE-888C2BAEEBDC}` |
| `ISpeechSynthesizerOptions2` | `{1CBEF60E-119C-4BED-B118-D250C3A25793}` |
| `ISpeechSynthesizerOptions3` | `{401ED877-902C-4814-A582-A5D0C0769FA8}` |
| `IVoiceInformation` | `{B127D6A4-1291-4604-AA9C-83134083352C}` |
| `IAsyncInfo` | `{00000036-0000-0000-C000-000000000046}` |
| `IRandomAccessStream` | `{905A0FE1-BC53-11DF-8C49-001E4FC686DA}` |
| `IInputStream` | `{905A0FE2-BC53-11DF-8C49-001E4FC686DA}` |
| `IBufferByteAccess` | `{905A0FEF-BC53-11DF-8C49-001E4FC686DA}` |
| `IBufferFactory` | `{71AF914D-C10F-484B-BC50-14BC623B3A27}` |

### Vtable slots

IUnknown occupies 0-2 and IInspectable 3-5 on every WinRT interface.

| Interface | Slot | Method |
|-----------|------|--------|
| `ISpeechSynthesizer` | 6 / 7 | `SynthesizeTextToStreamAsync` / `SynthesizeSsmlToStreamAsync` |
| | 8 / 9 | `put_Voice` / `get_Voice` |
| `ISpeechSynthesizer2` | 6 | `get_Options` |
| `ISpeechSynthesizerOptions2` | 6,7 / 8,9 / 10,11 | AudioPitch / AudioVolume / SpeakingRate (get, put) |
| `ISpeechSynthesizerOptions3` | 6 / 8 | AppendedSilence / PunctuationSilence (get) |
| `ISpeechSynthesizerStatics` | 6 | `get_AllVoices` |
| `IVectorView<T>` | 6 / 7 | `GetAt` / `get_Size` |
| `IVoiceInformation` | 6 / 7 / 8 | DisplayName / Id / Language |
| `IAsyncInfo` | 7 | `get_Status` |
| `IAsyncOperation<T>` | 8 | `GetResults` |
| `IAsyncOperationWithProgress<T,P>` | **10** | `GetResults` - Progress comes first |
| `IInputStream` | 6 | `ReadAsync` |
| `IRandomAccessStream` | 6 | `get_Size` |
| `IBuffer` | 7 | `get_Length` |
| `IBufferByteAccess` | 3 | `Buffer` (raw pointer) |

Generic interfaces such as `IAsyncOperation<T>` have IIDs computed by hashing their type
arguments, which is impractical here. They are called by slot instead, and completion is
awaited through the non-generic `IAsyncInfo`, whose IID is fixed.

### The two behavioural surprises

**`Options.SpeakingRate` does nothing.** It stores and reads back correctly but does not
change the audio. Verified with unique text per call to rule out caching:

```
SpeakingRate 0.5 -> 108046 bytes
SpeakingRate 1.0 -> 108046 bytes
SpeakingRate 3.0 -> 108046 bytes
```

Note also that all three of AudioPitch, AudioVolume and SpeakingRate default to `1.0`, so
they cannot be told apart by reading them. They can be told apart by range: setting `3.0`
on AudioPitch (valid 0.0-2.0) makes synthesis fail.

**SSML prosody rate saturates at 200%**, measured against `rate="100%"`:

```
 10% -> 0.58x     150% -> 1.24x      300% -> 1.54x
 50% -> 0.81x     200% -> 1.54x     1000% -> 1.54x
100% -> 1.00x
```

Named values reach further in both directions: `x-slow` 0.17x, `slow` 0.33x,
`medium` 0.50x, `fast` 0.81x, `x-fast` 1.54x.

### Output format

16 kHz, mono, 16-bit, delivered as a complete RIFF WAVE - so parse it with the `wave`
module rather than assuming a header length.

### Voice IDs match NVDA's

```
Microsoft David   HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech_OneCore\Voices\Tokens\MSTTS_V110_enUS_DavidM
Microsoft Zira    ...MSTTS_V110_enUS_ZiraM
Microsoft Mark    ...MSTTS_V110_enUS_MarkM
```

Same list, same identifiers, so matching the main synth's voice is a direct string
comparison.

---

## 6. What is missing from NVDA's own helper DLL

The export table of `nvdaHelperLocalWin10.dll` in NVDA 2026.1.1 contains twelve
`ocSpeech_*` functions and **no punctuation functions at all**:

```
getCurrentVoiceId  getCurrentVoiceLanguage  getPitch  getRate  getVoices  getVolume
initialize  setPitch  setRate  setVoice  setVolume  speak
supportsProsodyOptions  terminate
```

`ocSpeech_supportsPunctuationSilence` exists in NVDA's current source but not in that
build. Probing for it unguarded raises `AttributeError` and, if that probe sits in the
same `try` as the rest of setup, silently takes down everything with it.

Going to WinRT directly avoids this entirely, and gains punctuation control that NVDA's
build cannot offer.

---

## Sources

- [nvaccess/nvda](https://github.com/nvaccess/nvda)
- [NVDA Developer Guide](https://download.nvaccess.org/documentation/developerGuide.html)
