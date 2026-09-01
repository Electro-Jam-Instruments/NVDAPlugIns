# Spatial Typing Feedback - how it works and why

What was built and the reasoning behind it. Where a decision looks strange, the cause is
usually that something behaved differently from how it is documented - those are recorded
with the measurement that settled them, because they each cost hours to find.

Verified API details are in [`research/nvda-and-winrt-reference.md`](research/nvda-and-winrt-reference.md).

---

## The problem

NVDA speaks everything through one voice, in one place. Typed characters, completed words
and typing errors all arrive on the same channel as whatever NVDA is reading, competing
for the same moment of attention. Typing quickly means the character echo collides with
the document.

People separate concurrent sound well when it differs in **position** and **level**. This
add-on moves the character echo out of the way.

## The three streams

| Stream | Contains | Voice | Default |
|--------|----------|-------|---------|
| **Main** | NVDA's speech, and completed words | NVDA's own | centre |
| **Chars** | Each typed character | ours | right, quieter |
| **Errors** | The typing-error alert | none - a wave file | left, quieter |

**Only the character echo needed its own voice.** An earlier version also re-synthesised
completed words. That was wrong: word echo already works through NVDA's voice at the
user's own rate and rate boost, and it was never the stream getting in the way. Handing it
back removed a whole class of problems, including a bug where the space that completed a
word cancelled the word itself.

---

## Positioning is channel volume

`nvwave.WavePlayer.setVolume(*, all, left, right)` is the only per-channel control NVDA
exposes, and it is enough, because we create our own players.

**Constant-power pan law, not linear balance.** Our audio is mono duplicated into both
channels, so the two are fully correlated and a centred signal is up to 6 dB louder than
the same signal panned to one side. Linear balance would make the centre stream boom next
to the sides, and would sound like a volume bug rather than a pan bug.

```python
theta = (pan + 50) / 100 * (math.pi / 2)
left, right = math.cos(theta), math.sin(theta)   # centre -> ~0.707 each
```

**Mono has to be upmixed first.** `setVolume` raises `E_INVALIDARG` on channel 1 of a mono
player, so players are always created with `channels=2` and each 16-bit sample is
duplicated into both. Done with `array.array` extended slice assignment, which runs in C:
NVDA bundles no numpy and Python 3.13 removed `audioop`, so the alternative is a
per-sample Python loop, which is far too slow.

**Channel volume is applied only when it changes.** Setting it before every utterance
produced audible bursts - WASAPI channel volume changes are not sample accurate, so
setting them repeatedly against playing audio is heard. Steady-state typing now sets it
zero times.

---

## The character voice is a Windows SpeechSynthesizer we activate ourselves

The same engine NVDA's OneCore synthesizer uses: the same class, the same voices, the same
voice IDs.

### Why not through NVDA

`nvdaHelper`'s ocSpeech is a process-wide singleton:

```cpp
OcSpeechState g_state;                       // one m_synth, one m_callback

void* activate(std::function<ocSpeech_CallbackT> cb) {
    if(!isTerminated()){
        LOG_ERROR(L"Unable to activate if not terminated.");
        return nullptr;
    }
```

NVDA's own driver holds that slot. Tested rather than assumed: `ocSpeech_initialize`
returned null every time.

### What is actually true

The singleton is NVDA's wrapper, not Windows. Line 29 of that same C++ file:

```cpp
using winrtSynth = winrt::Windows::Media::SpeechSynthesis::SpeechSynthesizer;
```

`SpeechSynthesizer` is a public, documented, freely instantiable WinRT class. We activate
our own via `RoActivateInstance` and depend on nothing of NVDA's for the voice.

That removed fragility rather than just being tidier: no DLL version coupling, no reload
problems, and it gained punctuation-pause control that NVDA's own build cannot offer.

**SAPI 5 remains as a fallback** where no Windows voices exist or activation fails. Its
interface is deliberately identical - letting the two drift apart once cost an evening of
silent keystrokes.

---

## Two things the synthesizer does not do as documented

### `SpeechSynthesizerOptions.SpeakingRate` has no effect

It stores a value and reads it back correctly, which is why it looks like it works. The
audio is unchanged: identical byte counts at 0.5x, 1.0x and 3.0x, with unique text each
time to rule out caching.

Rate, pitch and volume are therefore set through **SSML prosody**, per utterance.

### SSML prosody rate saturates at 200%

Measured against `rate="100%"`:

| attribute | speed |
|-----------|-------|
| `10%` | 0.58x |
| `100%` | 1.00x |
| `200%` | **1.54x** |
| `300%`, `500%`, `1000%` | 1.54x - byte-identical audio |

Anything above 200% silently pins the voice at maximum, which sounds exactly like a broken
mapping. Speed maps onto **10%-200%** with rate boost on, and a narrower **50%-120%** with
it off for fine control near normal.

---

## Typing echo is intercepted at `speech.speakTypedCharacters`

NVDA decides both character and word echo in that one function, and it owns the typed-word
buffer.

**We do not reimplement its logic.** We run NVDA's own function and temporarily swap the
two functions it produces output through:

```python
speechModule.speakText = self._captureWord      # completed words -> back to NVDA
speechModule.speakSpelling = self._captureChar  # characters -> our voice
```

NVDA keeps the typed-word buffer, protected-field handling, the suppression counter and
terminal behaviour - state not all visible from outside. We take only what it decided to
say, and the swap lasts one synchronous call on the main thread.

This is a monkeypatch and the most version-fragile part of the add-on. It is guarded and
fails safe: if the shape has changed, NVDA behaves exactly as before.

**Characters interrupt characters.** Synthesis takes longer than a fast typist takes to
reach the next key, so a plain queue drifts further behind with every keystroke and sounds
like a voice set too slow. Interruption is kind-aware - a character must never cancel a
word queued behind it.

---

## The error alert uses a supported extension point

`nvwave.decide_playWaveFile` lets us cancel NVDA's playback and play the same file
ourselves, positioned. No patching.

NVDA reports spelling errors three ways and only one is a typing event:

| When | How | Ours |
|------|-----|------|
| **Typing** a word-ending character | `playWaveFile(textError.wav)`, `isSpeechWaveFileCommand=False` | **yes** |
| Reading over marked text | same wav, inside a speech sequence, so the flag is `True` | no |
| Reading over marked text | NVDA speaks "spelling error" | no |

The last two annotate text the main voice is reading; moving them would separate the
annotation from the words it describes. `isSpeechWaveFileCommand` distinguishes them.

**No words are spoken for errors.** The typing alert is sound-only in NVDA, so that stream
has no voice, speed or pitch - volume and pan are its only meaningful controls. Its volume
follows NVDA's own sound volume by default, mirroring `nvwave`'s rule including
`soundVolumeFollowsVoice`, and remains overridable.

---

## Sound Split is mutually exclusive with this add-on

Sound Split sets channel volume on the whole NVDA **process audio session**, one level
above our per-player volumes, so the two multiply. Rather than reason about which
combinations survive, the rule is one line: if Sound Split is on, this add-on stays
dormant and says so. It never changes the user's Sound Split setting.

---

## Tuning happens in the settings ring

Reachable with keys the user already has - `control+NVDA+arrows`. No new gestures.

The ring is a single swappable object: every ring script goes through
`globalVars.settingsRing`, and `synthDriverHandler` updates the *existing* object on synth
change rather than replacing it, so a subclass survives.

**One selector, one set of controls.** Pick a stream, then adjust it:

```
stream -> voice -> speed -> rate boost -> pitch -> volume -> punctuation -> pan
```

Our slots **replace** NVDA's rather than being appended. Appending meant Voice, Speed and
Volume each appeared twice - once acting on the main synth, once on the selected stream -
identical to hear, different in effect. NVDA's own settings remain in its Voice settings
dialog and return if the add-on is disabled.

Rules that came out of actually using it:

- **Slots are named after the stream** - "characters speed", not "spatial speed". The
  selected stream is restated at every step rather than having to be remembered.
- **Changes are spoken by the stream being tuned.** Adjusting the character voice while
  listening to the main voice - different voice, different place, different speed - tells
  you nothing about what you are changing. The announcement *is* the sample; a separate
  preview word was tried and removed, because it played the old settings immediately
  before the announcement played the new ones.
- **Inapplicable slots say "not available"** rather than disappearing, in the selected
  stream's voice. A control that vanishes and reappears as the selector moves is more
  disorienting than one that says it does not apply.
- **Settings that follow something else say so, and stay reachable.** The error volume's
  "follow NVDA" sits one step below zero, so overriding it is not a one-way door.

Settings save when you **leave the ring** - the first gesture that is not a ring command -
and on teardown. No timer: they are committed at a moment the user chose.

---

## Positioning is pan and volume only

No interaural delay, no HRTF, no bundled audio library.

3D is wanted later and nothing here blocks it: we own the PCM buffer, so a future version
can transform it however it likes. The route would be OpenAL Soft with `ALC_SOFT_loopback`,
rendering HRTF into a buffer we supply so we stay on NVDA's audio path. Cost: a native DLL
in two architectures, and megabytes of package. Generic HRTF gives good azimuth, reasonable
distance, poor front/back and weak elevation - worth knowing before promising anything.

---

## Deliberately not done

- **Application autocomplete.** "Word completion" here means NVDA's word echo, not
  suggestion lists. Autocomplete has no labelled hook - `filter_speechSequence` carries no
  provenance - so intercepting it means guessing, and guessing wrong in the dangerous
  direction swallows speech the user needed.
- **Panning the main voice.** We do not own NVDA's player. The slot says "not available"
  rather than pretending.
- **The caps-lock beep.** Cheap when wanted - `tones.beep` already takes `left` and
  `right` - but out of scope.
