# Spatial Typing Feedback - Architecture Decisions

All NVDA API facts below were verified against `nvaccess/nvda` master source
(see `docs/research/nvda-audio-and-typing-hooks.md` for file/line evidence).

---

## Decision D1 - GlobalPlugin, not AppModule

**Decision:** `addon/globalPlugins/spatialTypingFeedback.py`.

**Why:** Typing echo and spelling alerts happen in every application. An AppModule
would only cover one. Matches `windows-dictation-silence`.

---

## Decision D2 - How do we pan audio at all?

**Decision:** Use `nvwave.WavePlayer.setVolume(left=..., right=...)`.

NVDA's WASAPI player exposes exactly the primitive we need:

```python
def setVolume(self, *, all: float | None = None,
              left: float | None = None, right: float | None = None):
    """Levels must be specified as a number between 0 and 1."""
```

So a panned stream is: create our own `nvwave.WavePlayer(channels=2, ...)`, call
`setVolume(left=0.0, right=1.0)`, then `feed()` it PCM.

**Rejected - NVDA Sound Split:** `config.conf["audio"]["soundSplitState"]` separates
*NVDA as a whole* from *other applications*. It cannot separate NVDA's own streams
from each other, which is the entire point of this add-on. The two features should
coexist; we must not fight Sound Split for control of the main player's channel
volumes.

**Rejected - audio ducking:** wrong axis entirely (volume over time, not position).

---

## Decision D3 - Where does the secondary voice come from?

**Decision:** Create a **second, independent Windows OneCore instance** via NVDA's own
`ocSpeech` interface, render to a WAV buffer, and play it through *our own*
`nvwave.WavePlayer` with per-channel volume.

This gives the secondary voice the same engine, the same installed voices and the same
**rate boost** as NVDA's main synthesizer.

### Correcting an earlier claim

An earlier draft of this document recommended SAPI 5 and asserted that OneCore "holds
module-level global state, so a second instance would corrupt the first". **That is
wrong.** It is true of eSpeak; it is not true of OneCore.

OneCore is **token-based**. `source/synthDrivers/oneCore.py`:

```python
self._callbackInst = ocSpeech_Callback(self._callback)
self._ocSpeechToken = HANDLE()
self._ocSpeechToken.value = self._dll.ocSpeech_initialize(self._callbackInst)
```

Every subsequent call takes that token as its first argument -
`ocSpeech_speak(token, text)`, `ocSpeech_setRate(token, rate)`,
`ocSpeech_setVoice(token, index)`, `ocSpeech_terminate(token)`. The handle *is* the
instance. Nothing about the interface implies a process-wide singleton.

### The driver already has the architecture we want

NVDA's OneCore driver does not hand audio to some private internal path. It receives
PCM in a callback and feeds a `WavePlayer` it owns:

```python
def _callback(self, bytes, len, markers):
    # This gets called in a background thread.
    stream = io.BytesIO(ctypes.string_at(bytes, WAVE_HEADER_LENGTH))
    wav = wave.open(stream, "r")
    self._maybeInitPlayer(wav)
    data = bytes + WAVE_HEADER_LENGTH
    ...
    self._player.feed(ctypes.c_void_p(data + prevPos), size=dataLen - prevPos)
```

So the plan is to do the same thing with our own token, our own callback and our own
player - one whose channel volumes we control:

```
NVDAHelper.getHelperLocalWin10Dll()
  -> ocSpeech_initialize(ourCallback)   -> our own token
  -> ocSpeech_setVoice / setRate / setPitch / setVolume (our token)
  -> ocSpeech_speak(ourToken, text)
  -> our callback receives WAV bytes (header is WAVE_HEADER_LENGTH = 46)
  -> upmix mono to stereo
  -> our WavePlayer(channels=2).setVolume(left=, right=) -> feed()
```

The DLL itself is shared and already loaded by NVDA; we only ask for a handle to it.
The token is what isolates us.

### Rate boost, replicated exactly

Rate boost is not a hidden OneCore feature - it is a range change in the driver's
percentage-to-parameter mapping:

```python
MIN_RATE = 0.5
DEFAULT_MAX_RATE = 1.5
BOOSTED_MAX_RATE = 6.0

def _set_rate(self, rate):
    maxRate = self.BOOSTED_MAX_RATE if self._rateBoost else self.DEFAULT_MAX_RATE
    rawRate = self._percentToParam(rate, self.MIN_RATE, maxRate)
    self._queuedSpeech.append((self._dll.ocSpeech_setRate, rawRate))

@classmethod
def _percentToParam(cls, percent, min, max):
    return float(percent) / 100 * (max - min) + min
```

So rate boost is: map 0-100 onto 0.5-6.0 instead of 0.5-1.5, and pass the result to
`ocSpeech_setRate`. Ten lines. The secondary voice gets the same ceiling the main voice
has.

**Gated on `ocSpeech_supportsProsodyOptions()`** (OneCore API > 5). When false, rate,
pitch and volume cannot be changed after initialization at all, and there is no rate
boost - for anyone, including NVDA's main synth. Detect it and disable the control
rather than silently ignoring the setting.

### The one real obstacle: mono to stereo

`_maybeInitPlayer` builds its player from the wave header:

```python
self._player = nvwave.WavePlayer(
    channels=wav.getnchannels(),   # OneCore returns mono
    ...
)
```

OneCore returns **mono**, and `setVolume(left=, right=)` raises `E_INVALIDARG` on
channel 1 of a mono player. So we cannot simply pan the buffer we are given.

We must construct our player with `channels=2` and **upmix the mono PCM to stereo
before feeding it** - duplicate each 16-bit sample into both channels - then let
`setVolume` do the panning. This is cheap (a `numpy` repeat or a slice assignment over
a `bytearray`) but it is mandatory, and it is the single most likely thing to be got
wrong. Doubling the buffer also doubles the bytes fed per utterance; confirm that does
not matter at character-echo rates.

### Voice selection - match the main voice by default

Enumerate with `ocSpeech_getVoices(token)` - the same list the user already sees in
NVDA's synthesizer settings.

**Default the secondary voice to the *same* voice as the main synth**, with matching
rate, rate boost, pitch and volume. Separation is positional, not timbral (R6).

An earlier draft of this document recommended defaulting to a *different* voice on the
grounds that audible difference was "half the point". That was wrong. One familiar voice
in two places is easier to listen to than two voices competing, and the user is on a
specific OneCore voice for a reason. The user may still override the secondary voice
independently (R5); matching is the default, not a constraint.

Practically this means reading the main synth's current settings at startup and on
`synthDriverHandler.synthChanged`, and mirroring them onto our token unless the user has
set an explicit override.

### Costs and risks

- **Two OneCore instances in one process is unverified.** The token design says it
  should work. It has to be proven before anything is built on it - this is now the
  primary question for Spike S1.
- Our callback runs on a **background thread**, as NVDA's does. All the usual care
  around cancellation and teardown applies; NVDA uses an `_earlyExitCB` flag for
  exactly this and we should copy the pattern.
- We still own our own speech queue: interruption, backlog under fast typing, cleanup.

### Fallback

If S1 shows a second OneCore token is not viable, fall back to **SAPI 5 over COM**
(`comtypes` -> `SAPI.SpVoice` -> `SpMemoryStream`), same buffer-and-pan architecture,
lower voice quality and no rate boost. This is a genuine downgrade for a user who
chose OneCore *for* the rate boost, so it is a fallback, not an equal option.

---

## Decision D4 - How do we intercept typing echo? (R1 **and** R3)

**This single decision covers both the character echo and the word echo**, because NVDA
decides both in the same function. R1 and R3 are two branches of one hook, not two
separate features.

**Chain, verified in source:**

```
keyboardHandler
  -> NVDAObject.event_typedCharacter(ch)        # NVDAObjects/__init__.py
       -> speech.speakTypedCharacters(ch)       # speech/speech.py:1437
            -> speakText(typedWord)             # R3: word echo,  speakTypedWords
            -> speakSpelling(realChar)          # R1: char echo,  speakTypedCharacters
```

`speakTypedCharacters` is where both echo modes are decided, and it owns the
`_curWordChars` buffer that accumulates the word in progress. It is re-exported as
`speech.speakTypedCharacters`, and every caller reaches it through that package
attribute.

**Decision:** wrap `speech.speakTypedCharacters` with our own function that reproduces
NVDA's decision logic, routes each branch to the secondary voice at its own pan
position, and does **not** call through to `speakText` / `speakSpelling`.

| Branch | NVDA call | Our routing |
|--------|-----------|-------------|
| Word completed | `speakText(typedWord)` | secondary voice, **centre** |
| Character typed | `speakSpelling(realChar)` | secondary voice, **right** |

**Both branches can fire on the same keystroke.** `FIRST_NONCONTROL_CHAR = " "`
(speech.py:1423), so pressing space both completes the word *and* echoes as a
character. With both echo settings on, one space produces "hello" in the centre and
"space" on the right simultaneously. That is correct behaviour and is exactly the
separation the add-on exists to provide - but it means our secondary voice needs to
handle two overlapping utterances, which the queue design in D3 must account for.

**Why wrapping and not an overlay class:** `chooseNVDAObjectOverlayClasses` is the
sanctioned extension point and would normally win. But the echo decision does not live
on the object - it lives in `speakTypedCharacters`, together with the word buffer we
must not duplicate. An overlay would have to reimplement the buffer and would still
race with `speech.clearTypedWordBuffer()` calls made elsewhere (terminals do this).
Wrapping the single function is the smaller, more honest hook.

**Constraints we must reproduce exactly:**

- `TypingEcho` is `OFF=0, EDIT_CONTROLS=1, ALWAYS=2` (`config/configFlags.py:55`).
  Note EDIT_CONTROLS is 1 and ALWAYS is 2 - easy to get backwards.
- `api.isTypingProtected()` must still substitute the protected character.
- `_speechState._suppressSpeakTypedCharactersNumber` suppression must be respected.
- Terminals queue characters instead of echoing them
  (`NVDAObjects/behaviors.py:622`); we must not double-speak there.

**Risk:** this is a monkeypatch on NVDA internals. It is version-fragile by
construction. Mitigation: guard the patch behind a version and signature check, log
loudly, and fall back to a no-op (NVDA behaves normally) if the shape has changed.
Never leave NVDA silent because our patch failed.

---

## Decision D5 - How do we intercept typing errors? (R2)

**Chain, verified in source** (`NVDAObjects/behaviors.py:306`):

```python
if (config.conf["documentFormatting"]["reportSpellingErrors2"] != OFF
    and config.conf["keyboard"]["alertForSpellingErrors"]
    and (ch.isspace() or (ch >= " " and ch not in "'\x7f" and not ch.isalpha()))):
    self._reportErrorInPreviousWord()
```

`_reportErrorInPreviousWord` waits 50ms (`core.callLater`) for the app to mark the
error, checks for an `invalid-spelling` format field, then plays:

```python
nvwave.playWaveFile(os.path.join(globalVars.appDir, "waves", "textError.wav"))
```

**Decision:** register a handler on the **`nvwave.decide_playWaveFile` extension
point**. When the filename is `textError.wav`, return `False` to suppress NVDA's
unpanned playback, and play the same wave ourselves through a `WavePlayer` panned left.

**Why this is the right hook:** it is a real, supported extension point (not a
monkeypatch), it is filename-addressable, and it leaves all of NVDA's detection logic -
including the 50ms delay and the format-field check - completely untouched. We are only
changing *where the sound comes from*, which is exactly the requirement.

### There are three error paths, and only one of them is ours

NVDA reports spelling and grammar errors in three different ways. They are not
interchangeable and only the first belongs in the Errors stream.

| # | When | How | Ours? |
|---|------|-----|-------|
| 1 | **While typing** - you finish a misspelled word | `playWaveFile(textError.wav)`, `isSpeechWaveFileCommand=False`. **Sound only - there is no speech path here at all.** | **Yes - pan left** |
| 2 | **While reading** over an error, SOUND bit set | Same wav, but emitted as `WaveFileCommand` *inside the speech sequence* -> `playWaveFile(..., isSpeechWaveFileCommand=True)` | No - leave it |
| 3 | **While reading** over an error, SPEECH bit set | NVDA literally speaks `"spelling error"` / `"grammar error"` / `"out of spelling error"` | No - leave it |

The setting is a **bitmask**, not a toggle (`config/configFlags.py:185`):

```python
class ReportSpellingErrors(DisplayStringIntFlag):
    OFF = 0b0
    SPEECH = 0b1
    SOUND = 0b10
    SPEECH_AND_SOUND = SPEECH | SOUND
    BRAILLE = 0b100
```

which is why `reportSpellingErrors2` is `integer(min=0, max=7)` - `SPEECH|SOUND|BRAILLE`.

**So yes, NVDA does speak errors - but not while you are typing.** It speaks them when
you read or navigate over text that is already marked.

### Why paths 2 and 3 must stay with the main voice

Both are *descriptions of text the main voice is currently reading*. The error sound and
the words "spelling error" arrive interleaved with the sentence they annotate. Moving
them to the left would divorce the annotation from the thing it annotates - the user
would hear "spelling error" on the left with no idea which word it referred to.

They belong exactly where they are: in the main stream, in sequence.

**`isSpeechWaveFileCommand` is how we tell them apart.** The `decide_playWaveFile`
extension point receives it, so our handler pans only when it is `False`:

```python
def _decide_playWaveFile(self, fileName, asynchronous, isSpeechWaveFileCommand, **kwargs):
    if isSpeechWaveFileCommand:
        return True          # part of a speech sequence - leave it alone
    if os.path.basename(fileName).lower() != "texterror.wav":
        return True
    ...                      # ours: play it panned left, return False
```

### The Errors stream needs no voice

This is the useful consequence. Streams 1 and 2 (characters, words) need the secondary
OneCore instance. The Errors stream is **a wave player and nothing else** - no synth, no
token, no upmix of synthesised audio, just a wav file read once and fed panned.

That is why P1 can ship before spike S1 resolves: the error alert does not depend on the
second OneCore token being viable at all.

### Optional: speak the typing alert instead of thumping

Since we will already have a panned voice, we could offer to *speak* the typing-time
alert - a short word like "error" in the secondary voice, panned left - as an
alternative to `textError.wav`. NVDA cannot do this today; path 1 is sound-only by
construction.

Not required by R2. Worth offering in the settings panel once P2 exists, because some
users read a word faster than they classify a thump.

### Scope: the spelling alert only

R2 covers NVDA's **typing-time spelling alert** and nothing else.

**Deferred to a later version:** the caps-lock beep
(`config.conf["keyboard"]["beepForLowercaseWithCapslock"]`, fired from
`NVDAObject.event_typedCharacter`). Cheap when we want it - `tones.beep` already accepts
`left=` and `right=`, so it needs none of the wave-player machinery here. Tracked in
`docs/TODOs/next-steps.md`.

**Out of scope:** NVDA's own internal error sound, and application-level input
validation.

---

## Decision D6 - Application autocomplete is out of scope

**Superseded.** An earlier draft of this document read requirement R3 as *application
autocomplete / suggestion lists* and designed a two-stage interception around
`InputFieldWithSuggestions.event_suggestionsOpened` and
`speech.extensions.filter_speechSequence`.

That was a misreading. R3 is NVDA's **word echo** - speaking the whole word when the
user finishes typing it - which is handled entirely by Decision D4.

The distinction matters, so it is recorded rather than deleted:

| | Word echo (R3, in scope) | Autocomplete (out of scope) |
|---|---|---|
| Trigger | User finishes typing a word | App offers a completion |
| NVDA path | `speakTypedCharacters` -> `speakText` | Ordinary object presentation |
| Setting | `keyboard.speakTypedWords` | `presentation.reportAutoSuggestionsWithSound` |
| Hook | Wrap one function (D4) | No labelled hook; must infer provenance |

Dropping autocomplete removes the riskiest part of the design. `filter_speechSequence`
receives a speech sequence with no indication of *why* it is being spoken, so
intercepting suggestion text would have meant guessing - and guessing wrong in the
dangerous direction means swallowing speech the user needed.

If autocomplete panning is wanted later it is a separate feature with a separate spike.
The background research is retained in
`docs/research/nvda-audio-and-typing-hooks.md` section 4.

---

## Decision D7 - Sound Split and this add-on are mutually exclusive

**Decision:** if this add-on is active, Sound Split is not. One or the other.

**Why:** Sound Split sets channel volume on the whole NVDA **process audio session**
(`audio/soundSplit.py`), one level above our per-player `setVolume`. The two multiply,
and the results are not worth reasoning about case by case.

**Behaviour:**
- Read `config.conf["audio"]["soundSplitState"]` at startup and on config change.
- If Sound Split is on, stay dormant - no panning, echo through NVDA's normal path - and
  say so once, naming the cause.
- Never change the user's Sound Split setting. Offer, do not act. Never call
  `_setSoundSplitState` or `_toggleSoundSplitState`; the underscore means private.

Two earlier drafts got this wrong in opposite directions - first treating it as a vague
advisory, then enumerating which Sound Split states happen to survive. Neither is worth
the complexity. The rule is one line: **our feature on means Sound Split off.**

---

## Decision D8 - Settings ring integration (R7)

**Decision:** Replace `globalVars.settingsRing` with a subclass of NVDA's
`SynthSettingsRing` that appends our two slots after the synth's own.

### Why this works

The ring is a single swappable object, and the gestures never touch the class:

```python
# globalCommands.py
def script_nextSynthSetting(self, gesture):
    nextSettingName = globalVars.settingsRing.next()
    ...
def script_increaseSynthSetting(self, gesture):
    settingValue = globalVars.settingsRing.increase()
```

Every ring script goes through `globalVars.settingsRing`. Swap that object and the
existing `Ctrl+NVDA+Arrow` gestures reach our slots with no gesture registration and no
patching of `globalCommands`.

The object survives synth changes, too (`synthDriverHandler.py:445`):

```python
# start or update the synthSettingsRing
if globalVars.settingsRing:
    globalVars.settingsRing.updateSupportedSettings(synth)
else:
    globalVars.settingsRing = SynthSettingsRing(synth)
```

Because the truthy branch calls `updateSupportedSettings` on the *existing* object, NVDA
keeps updating ours rather than replacing it. We override that method to call `super()`
and then append our slots, so they persist across synth switches and land after the
synth's own settings.

### Why not the alternatives

- **A separate ring on its own gestures.** Simpler and safer, but it is a second thing
  to learn and a second set of keys to remember. R7 explicitly asks for the ring the
  user already uses.
- **Injecting into `synth.supportedSettings`.** The ring builds from that list, so it
  would work - but `SynthSetting._set_value` then writes
  `config.conf["speech"][synth.name][setting.id]`, putting our settings inside NVDA's
  per-synth speech config, against a spec that does not declare them. Wrong place,
  likely to fail validation.

### The two slots

Our settings subclass `SynthSetting` and override `_set_value` to persist into our own
config section rather than `config.conf["speech"][synthName]`.

**Slot 1 - stream selector.** A string-valued slot cycling **Main / Chars /
Errors** - the three streams. Changing it makes no sound change; it retargets slot 2.

Selecting Main adjusts the main voice and word echo together, which is what the user
expects from three streams. If they ever need to diverge, adding a fourth value is a
one-line change - the audio paths are already separate.

**Slot 2 - Pan.** A `NumericDriverSetting`, range **-50 (left) to +50 (right), 0 =
centre** - the familiar balance convention. Suggested `normalStep=5`, `largeStep=25`, so
one large step moves a full quadrant and five normal steps reach an extreme. Negative
`minVal` is fine; the ring only does clamped arithmetic
(`min(self.max, self.value + self.step)`).

Report the value as a position, not a bare number. "Left 50", "Centre", "Right 25" is
usable at boosted rate; "minus fifty" is not.

### Starting values - three streams: Main, Chars, Errors

| Stream | Contains | Pan | Volume |
|--------|----------|-----|--------|
| **Main** | Main voice + completed words | 0 | 100% |
| **Chars** | Typed characters | +35 | 50% |
| **Errors** | Typing errors | -35 | 50% |

**Three positions, not four.** Main speech and word echo share the centre. They remain
separate audio paths in code - NVDA's synth and our secondary voice - so a future version
could split them, but they are one place as far as the user is concerned.

**Symmetric sides.** Characters and errors sit the same distance out at the same level.
The centre stream is what you are actually listening to; the sides are peripheral by
position *and* by volume.

**Volume is per stream, not global.** Halving the side streams does as much work as
panning them - arguably more.

Neither side sits at a hard extreme. Full -50/+50 is available but is not the starting
point.

**These are starting values for the spike, not persisted defaults.** Do not build the
config layer around them until they have been heard - see D11.

**A volume slot will be needed too.** Per-stream volume is part of the spec, so the ring
eventually wants Stream / Pan / Volume. Pan first; volume once the values have settled.

**Slot 4 (recommended, not required) - multi-stream on/off.** R8 asks for the master
toggle in the settings panel. NVDA already supports boolean ring slots
(`BooleanSynthSetting`, reported as "on"/"off"), so exposing it in the ring as well is
nearly free and makes A/B comparison a single keystroke.

### Naming the selector slot

"Voice" cannot be used - NVDA's ring already has a Voice slot that selects the
synthesizer voice, and two slots called Voice in one ring is worse than a mediocre name.

| Candidate | Assessment |
|-----------|------------|
| **Stream** | **Recommended.** One syllable, no collision, survives boosted rate |
| Channel | Collides with stereo channels, which is what Pan adjusts |
| Output | Collides with NVDA's audio "output device" |
| Target | Vague when spoken alone in a ring |
| Feedback stream | Unambiguous but long to hear on every ring step |

### The trap that must be designed around

NVDA's built-in Rate, Pitch, Volume and Voice slots act on the **main synthesizer**,
always. If the user sets Stream to "Characters" and then steps to Rate, they will change
the main voice's rate, not the character echo's.

The Stream slot scopes our slots only. Mitigations:
- Name our slots so ownership is audible. If per-stream rate is added later it must be
  "Echo rate", never "Rate".
- Speak the selector change with its context: "Stream: characters", not "characters".

### Pan law - do not use plain linear balance

Our source is **mono duplicated into both channels** (D3), so the two channels are fully
correlated. A correlated signal at equal level in both channels is up to **6 dB louder**
than the same signal panned hard to one side. Linear balance would therefore make the
centre-panned main voice noticeably louder than the hard-panned character echo - and the
user would reach for the volume control to fix something that is actually a pan-law bug.

Use a constant-power law instead, so perceived loudness stays flat across the sweep:

```python
import math

def panToChannels(pan: int) -> tuple[float, float]:
    """pan: -50 (hard left) .. 0 (centre) .. +50 (hard right)"""
    theta = (pan + 50) / 100 * (math.pi / 2)
    return math.cos(theta), math.sin(theta)   # centre -> ~0.707 each
```

Verify by ear during S1 - this is a judgement call about loudness, not a calculation.

### Do not let NVDA reset our channel volumes

`nvwave.WavePlayer` re-applies a flat volume on `open()` and on `stop()`:

```python
def open(self):
    ...
    self._setVolumeFromConfig()

def stop(self):
    ...
    self._setVolumeFromConfig()

def _setVolumeFromConfig(self):
    if self._purpose is not AudioPurpose.SOUNDS:
        return
    volume = config.conf["audio"]["soundVolume"]
    ...
    self.setVolume(all=volume / 100)
```

`setVolume(all=...)` sets both channels equal - it **wipes any panning**.

The guard saves us for speech: it returns early unless `_purpose is
AudioPurpose.SOUNDS`, and speech players default to `AudioPurpose.SPEECH`. So the echo
streams are safe.

**The error alert player is not.** If we create it with `purpose=AudioPurpose.SOUNDS`
(which is semantically correct for a wave file), its pan is reset every open and every
stop, and the alert plays centred. Either:
- re-apply `setVolume(left=, right=)` immediately before each `feed()`, and fold
  `config.conf["audio"]["soundVolume"]` into our own coefficients so the user's sound
  volume setting is still honoured; or
- use `AudioPurpose.SPEECH` for that player and accept the semantic mismatch.

Prefer the first. Losing the user's sound volume setting to fix a panning bug is a bad
trade.

### Scope

R7 is phase P4, after the features work. It adjusts settings that must already exist.

---

## Decision D9 - Panning the main voice (R7)

R7 says the main voice is a stream like any other, defaulting to centre and adjustable
from the ring. That is a bigger ask than it sounds, because we do not own the main
synth's player.

### Centre is free; anything else is not

The main voice already plays at equal volume in both channels, which *is* centre. So the
default position costs nothing - the feature only bites when the user moves the main
voice off centre.

### The hook

NVDA's OneCore driver holds its player on a plain attribute:

```python
def _maybeInitPlayer(self, wav):
    ...
    self._player = nvwave.WavePlayer(
        channels=wav.getnchannels(),
        samplesPerSec=samplesPerSec,
        bitsPerSample=bytesPerSample * 8,
        outputDevice=config.conf["audio"]["outputDevice"],
    )
```

So `synthDriverHandler.getSynth()._player.setVolume(left=, right=)` pans the main voice.

**But the main synth's player is created with `channels=wav.getnchannels()`, and OneCore
is mono.** Same problem as D3: `setVolume` raises `E_INVALIDARG` on channel 1 of a mono
player. We cannot upmix the main synth's audio - it is not ours.

**This is the open question for spike S1.** Possibilities, in order of preference:

1. The output device is stereo and WASAPI presents two channels regardless of the source
   being mono, in which case `setVolume` works and this is easy. Test first.
2. If not, the main voice can only be centre, and R7's "main is adjustable" degrades to
   "main is centre, the others move around it". Given centre is the default and the
   sensible position for the primary stream, that is a mild loss - but the user must be
   told rather than finding a control that silently does nothing.
3. Replacing the main synth's player wholesale is not on the table. Too invasive.

### Constraints if it does work

- `_player` is **`None` until the synth first speaks.** Apply lazily, not at startup.
- `_maybeInitPlayer` **replaces** the player whenever the sample rate changes, silently
  discarding our pan. Re-apply on `speech.extensions.pre_speech` (a supported extension
  point) rather than assuming it sticks.
- `_player` is a private, driver-specific attribute. `getattr(synth, "_player", None)`,
  guard everything, and fail to centre rather than raising.
- Purpose is `AudioPurpose.SPEECH`, so `_setVolumeFromConfig` no-ops and will not wipe
  the pan (see D8).

### Restore on disable

R8's master off switch must restore the main player to `setVolume(all=1.0)`. Leaving the
user's main voice panned after they turned the feature off would be the worst failure
this add-on could produce.

---

## Decision D10 - Pan and volume only for now; 3D HRTF kept as a future option

**Decision:** positioning is **`WavePlayer.setVolume(left=, right=)` and volume. Nothing
else.** No interaural delay, no HRTF, no bundled audio library.

**Why:** it is enough to ship, it has no dependencies, and it keeps the add-on small.
Every other decision in this document already assumes it.

### Kept as a future option: 3D head-related transfer functions

Not rejected - deferred, and wanted. Recorded so it stays cheap to pick up later.

The enabling fact is that **we own the PCM buffer** (D3). NVDA only ever sees finished
stereo samples, so a future version can transform them however it likes without needing
anything new from NVDA. `setVolume` is the only positioning NVDA *offers*; it was never
a ceiling.

The route, when we want it: **OpenAL Soft** with the `ALC_SOFT_loopback` extension, so
HRTF renders into a buffer we supply and we stay on NVDA's audio path - keeping output
device, ducking and sound volume. Position becomes `alSource3f(source, AL_POSITION, x, y, z)`.

What it would cost, and what it would actually deliver, is in
`docs/research/spatial-audio-options.md`. The headlines:

- A native DLL in **both x86 and x64** builds (NVDA master is x86_64; 2024.1 and 2025.1
  are x86), LGPL compliance, and megabytes instead of kilobytes.
- No numpy, no scipy, and no `audioop` on Python 3.13 - so rolling our own convolution is
  not viable. It is OpenAL or nothing.
- Azimuth works well; distance is reasonable; **front/back is unreliable and elevation is
  weak** with generic HRTF. A ring around the head plus near/far is the honest promise.
- Headphones required. A single earbud breaks binaural rendering entirely.

Nothing in the current design blocks this. Keeping pan-and-volume as the only positioning
today does not paint us into a corner.

---

## Decision D11 - Do not build the settings layer until the values have been heard

**Decision:** S1 and P1/P2 use **hardcoded** stream values. No config schema, no
persistence, no settings panel until the sound has been judged by ear.

**Why:** the numbers in D8 - centre, +35, 50% volume - are a considered starting point,
not a result. Whether the character echo at half volume is still intelligible under a
boosted main voice is a listening question, and there is no way to answer it in advance.

Building the config layer first would mean writing a schema, migration and a panel around
values that are likely to move, and then rewriting all three. Worse, persisted defaults
acquire authority they have not earned - a value in a settings file looks decided.

**Sequence:**
1. Spike and early phases: constants in one module, easy to edit and rebuild.
2. Listen. Adjust. Repeat.
3. Only once the values stop moving, freeze them as defaults and build P3's panel and
   P4's ring slots around them.

This is why P3 sits after P2 in the phase table rather than alongside it.

---

## Phasing

| Phase | Delivers | Requirement |
|-------|----------|-------------|
| S1 | Spike: **second OneCore token**, mono-to-stereo upmix, panning, rate boost | D2, D3, R6 |
| P1 | Panned error alert (lowest risk, useful on its own) | R2 |
| P2 | Character echo right **and** word echo centre, secondary OneCore voice | R1, R3, R6 |
| P3 | Settings panel, incl. multi-stream master on/off | R5, R6, R8 |
| P4 | Settings ring slots: Stream selector and Pan, incl. main voice | R7, D9 |
| - | *Future option, not scheduled:* 3D HRTF via OpenAL Soft | R9, D10 |

P1 first, deliberately: it depends only on a supported extension point, it is
independently useful, and it proves the audio path before we touch the speech pipeline.

P2 delivers R1 and R3 together because they are one hook (D4). Splitting them would
mean patching `speakTypedCharacters` twice.

**S1 gates everything that uses a voice.** P1 does not - the error alert is a wave file,
not speech - so P1 can proceed in parallel with the spike. P2 and P3 cannot start until
S1 confirms a second OneCore token is viable, because the fallback (SAPI 5, no rate
boost) would change what those phases deliver.
