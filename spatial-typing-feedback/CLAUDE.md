# Spatial Typing Feedback Plugin - Development Instructions

## Critical Patterns

These patterns are MANDATORY for this plugin.

### GlobalPlugin (not AppModule)
```python
from globalPluginHandler import GlobalPlugin

class GlobalPlugin(GlobalPlugin):
    pass
```
**Why:** Typing echo and spelling alerts happen in every application.

See: `docs/architecture-decisions.md` Decision D1

### Panning is `WavePlayer.setVolume`, keyword-only
```python
player = nvwave.WavePlayer(channels=2, samplesPerSec=22050, bitsPerSample=16)
player.setVolume(left=0.0, right=1.0)   # hard right
```
**Why:** This is the only channel-level volume primitive NVDA exposes. Arguments are
keyword-only, levels are 0.0-1.0, and `all=` cannot be combined with `left=`/`right=`.
Mono devices raise `E_INVALIDARG` on channel 1 - degrade, do not crash.

See: `docs/architecture-decisions.md` Decision D2

### Secondary voice is a second OneCore token, not a second synth driver
```python
dll = NVDAHelper.getHelperLocalWin10Dll()
dll.ocSpeech_initialize.restype = HANDLE
token = HANDLE()
token.value = dll.ocSpeech_initialize(ourCallbackInst)   # our own instance
dll.ocSpeech_speak(token, text)                          # every call takes the token
```
**Why:** OneCore is token-based, so instances are independent. Do NOT instantiate
`OneCoreSynthDriver` - go to the `ocSpeech` layer directly, so we own the callback and
therefore the `WavePlayer` we pan. eSpeak *is* global-state bound; OneCore is not.

See: `docs/architecture-decisions.md` Decision D3

### ALWAYS upmix mono to stereo before panning
OneCore returns **mono**. `setVolume(left=, right=)` raises `E_INVALIDARG` on channel 1
of a mono player. Build the player with `channels=2` and duplicate each 16-bit sample
into both channels before `feed()`.

**Why:** this is the single most likely thing to get wrong, and it fails as a confusing
COM error rather than as silence.

See: `docs/architecture-decisions.md` Decision D3

### Rate boost is just a range change
```python
MIN_RATE, DEFAULT_MAX_RATE, BOOSTED_MAX_RATE = 0.5, 1.5, 6.0
maxRate = BOOSTED_MAX_RATE if rateBoost else DEFAULT_MAX_RATE
rawRate = float(percent) / 100 * (maxRate - MIN_RATE) + MIN_RATE
dll.ocSpeech_setRate(token, rawRate)
```
**Why:** the user is on OneCore *for* the rate boost (R6). A secondary voice without it
is unusable alongside a boosted main voice. Gate on
`ocSpeech_supportsProsodyOptions()`.

See: `docs/architecture-decisions.md` Decision D3

### Prefer extension points over monkeypatching
```python
nvwave.decide_playWaveFile.register(self._decide_playWaveFile)   # supported
speech.speakTypedCharacters = self._wrapped                      # LAST RESORT
```
**Why:** Only the typing echo path (D4) has no supported hook. The error sound does.
Any monkeypatch must be version-guarded and must fail safe - if it cannot be applied,
NVDA behaves normally.

See: `docs/architecture-decisions.md` Decisions D4, D5

### One hook covers both character echo and word echo
`speech.speakTypedCharacters` decides *both*: `speakSpelling(realChar)` for characters
and `speakText(typedWord)` for completed words. Patch it once, route the two branches
to different pan positions.

```python
# character -> right      word -> centre
```

**Why:** R1 and R3 are two branches of one function, not two features. Patching twice
would be wrong. Note that one space fires both branches, so two utterances can overlap.

See: `docs/architecture-decisions.md` Decision D4

### Application autocomplete is NOT this add-on
"Word completion" in the requirements means NVDA's word echo, not suggestion lists.
Do not build suggestion interception.

See: `docs/architecture-decisions.md` Decision D6

### Settings ring: swap the object, do not patch the scripts
```python
globalVars.settingsRing = OurRing(synth)   # subclass of SynthSettingsRing
# override updateSupportedSettings: call super(), then append our slots
```
**Why:** every ring gesture goes through `globalVars.settingsRing`, and
`synthDriverHandler` calls `updateSupportedSettings` on the *existing* object, so ours
survives synth changes. Note `self.settings` can be `None`, not `[]`.

Our slot classes must override `_set_value` - NVDA's writes to
`config.conf["speech"][synthName]`, which is the wrong place for our settings.

See: `docs/architecture-decisions.md` Decision D8

### Positioning is pan and volume ONLY
No interaural delay, no HRTF, no bundled audio library. 3D is a documented future option
(D10), not current scope. Do not build it.

If you ever do touch audio maths: NVDA has no numpy, no scipy, and Python 3.13 removed
`audioop`. Use `array.array` extended slice assignment (runs in C); never a per-sample
Python loop.

See: `docs/architecture-decisions.md` Decision D10

### Pan scale is -50 / 0 / +50
Left / centre / right, the familiar balance convention. NOT -100 to +100.

### Three streams: Main, Chars, Errors - pan AND volume, per stream
| Stream | Contains | Pan | Volume |
|--------|----------|-----|--------|
| Main | main voice + completed words | 0 | 100% |
| Chars | typed characters | +35 | **50%** |
| Errors | typing errors | -35 | **50%** |

Three positions, not four - main and word echo share centre on purpose. They stay
separate audio paths in code, so they can be split later if needed.

Sides are symmetric: same distance out, same level. Volume is per stream, not one global
secondary-voice level.

### Hardcode these values; do NOT build the config layer yet
Constants live in `_streams.py`. No schema, no persistence, no panel until the sound has
been judged by ear and the numbers stop moving. A value in a settings file looks decided
when it is not.

To try different values: edit `_streams.py`, then `NVDA+control+F3` to reload plugins.

See: `docs/architecture-decisions.md` Decision D11

### Pan law: constant power, never linear balance
```python
theta = (pan + 50) / 100 * (math.pi / 2)
player.setVolume(left=math.cos(theta), right=math.sin(theta))   # centre -> ~0.707
```
**Why:** our source is mono duplicated into both channels, so it is fully correlated.
Linear balance makes centre up to 6 dB louder than hard-panned - completed words would
boom next to the character echo, and it would read as a volume bug rather than a pan
bug.

See: `docs/architecture-decisions.md` Decision D8

### The error alert player will lose its pan unless you re-apply it
`WavePlayer.open()` and `.stop()` call `_setVolumeFromConfig()`, which calls
`setVolume(all=...)` and wipes panning - but **only** when
`purpose is AudioPurpose.SOUNDS`. Speech players are safe; the error alert is not.
Re-apply pan before each `feed()`, and fold `config.conf["audio"]["soundVolume"]` into
the coefficients so the user's setting still applies.

See: `docs/architecture-decisions.md` Decision D8

### Secondary voice MATCHES the main voice
Same OneCore voice, same rate, same rate boost, same pitch. Separation is positional,
not timbral. Do not default to a different voice.

See: `docs/architecture-decisions.md` Decision D3, requirement R6

### Sound Split on means this add-on stays dormant
```python
if config.conf["audio"]["soundSplitState"] != SoundSplitState.OFF:
    # no panning, echo through NVDA's normal path, say so once
```
**Why:** Sound Split sets channel volume on the whole NVDA process session, above our
per-player volumes. Mutually exclusive - one or the other. Never call
`_setSoundSplitState`.

See: `docs/architecture-decisions.md` Decision D7

### Restore the main player on disable
If we pan the main voice (D9), the master off switch MUST call `setVolume(all=1.0)` on
it. Leaving the user's main voice panned after they turned the feature off is the worst
failure this add-on could produce.

See: `docs/architecture-decisions.md` Decision D9

## Documentation

| Need | Location |
|------|----------|
| User requirements | `docs/user-requirements.md` |
| Why decisions were made | `docs/architecture-decisions.md` |
| Exact NVDA APIs, with source quotes | `docs/research/nvda-audio-and-typing-hooks.md` |
| What to build next | Phase table at end of `docs/architecture-decisions.md` |
| Why settings are not built yet | `docs/architecture-decisions.md` Decision D11 |

## Key Technical Facts

- **Panning:** `nvwave.WavePlayer.setVolume(*, all=, left=, right=)` - `source/nvwave.py`
- **Echo:** `speech.speakTypedCharacters(ch)` - `source/speech/speech.py`, handles
  characters (`speakSpelling`) and completed words (`speakText`)
- **Word completion trigger:** a character whose unicode category is not L, M or N,
  with a non-empty word buffer. `FIRST_NONCONTROL_CHAR = " "`, so space fires both branches
- **Echo modes:** `TypingEcho.OFF=0, EDIT_CONTROLS=1, ALWAYS=2` - note the order
- **Error alert:** `waves/textError.wav` via `nvwave.playWaveFile`, gated on
  `keyboard.alertForSpellingErrors` and `documentFormatting.reportSpellingErrors2`
- **Wave interception:** `nvwave.decide_playWaveFile.decide(fileName=, asynchronous=, isSpeechWaveFileCommand=)`
- **Secondary voice:** `ocSpeech_initialize(callback) -> HANDLE` token, then
  `ocSpeech_speak/setVoice/setRate/setPitch/setVolume/getVoices/terminate(token, ...)`
  - `source/synthDrivers/oneCore.py`
- **WAV header from the ocSpeech callback:** `WAVE_HEADER_LENGTH = 46`; PCM starts at
  `bytes + 46`. Callback runs on a **background thread** - copy NVDA's `_earlyExitCB`
  teardown pattern
- **Rate boost:** `MIN_RATE=0.5, DEFAULT_MAX_RATE=1.5, BOOSTED_MAX_RATE=6.0`, gated on
  `ocSpeech_supportsProsodyOptions()`
- **Settings ring:** `globalVars.settingsRing`, rebuilt via
  `updateSupportedSettings(synth)` from `synth.supportedSettings` filtered on
  `availableInSettingsRing` - `source/synthSettingsRing.py`
- **Ring gestures (already bound by NVDA):** `Ctrl+NVDA+Left/Right` moves between
  settings, `Ctrl+NVDA+Up/Down` changes the value. We register no gestures for R7
- **Streams:** Main (main voice + words) / Chars / Errors. Values: 0 at 100%,
  +35 at 50%, -35 at 50%
- **Main voice player:** `synthDriverHandler.getSynth()._player` - `None` until first
  speech, replaced on sample-rate change, private and driver-specific. Guard everything
- **NVDA 2026.1** renamed `WasapiWavePlayer` to `WavePlayer` - version guards must handle both

## Deployment

Tag format: `spatial-typing-feedback-vX.X.X-beta` (plugin prefix REQUIRED)

Use `/deploy` command or see `../.claude/commands/deploy.md`

## Code Layout

| File | Holds |
|------|-------|
| `__init__.py` | GlobalPlugin: lifecycle, Sound Split guard, voice matching, scripts |
| `_streams.py` | The three streams and their hardcoded pan/volume values (D11) |
| `_audio.py` | `PannedPlayer`: pan law, mono-to-stereo upmix, channel gains |
| `_voice.py` | `OneCoreVoice`: second `ocSpeech` token, rate boost, audio callback |
| `_echo.py` | `EchoInterceptor`: wraps `speech.speakTypedCharacters` |
| `_errors.py` | `ErrorAlertInterceptor`: `decide_playWaveFile` hook |

### The echo interception does NOT reimplement NVDA's logic
It runs NVDA's own `speakTypedCharacters` and temporarily swaps
`speech.speech.speakText` / `speakSpelling` to capture the output. NVDA keeps the
typed-word buffer, protected-field handling, suppression counter and terminal
behaviour. Do not replace this with a reimplementation - the state is not all visible
from outside.

## Current Status

v0.1.0-beta - All three streams working. Values hardcoded pending listening (D11).
Settings panel (P3) and settings ring slots (P4) not built yet.
