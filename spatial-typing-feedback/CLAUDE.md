# Spatial Typing Feedback - development instructions

Design and reasoning: `docs/architecture.md`. Verified API details, IIDs and vtable
slots: `docs/research/nvda-and-winrt-reference.md`.

## Things that are not what they appear

Each of these cost hours. Read them before touching the voice or the audio.

### `SpeechSynthesizerOptions.SpeakingRate` does nothing
It stores a value and reads it back correctly, so it looks like it works. The audio is
unchanged. Rate, pitch and volume are set through **SSML prosody**, per utterance.

### SSML prosody rate saturates at 200%
`300%`, `500%` and `1000%` produce byte-identical audio to `200%`. Sending a larger
number silently pins the voice at maximum, which sounds like a broken mapping.

```
boost on:  10%-200%      boost off: 50%-120%  (fine control near normal)
```

### `setVolume` on a mono player raises `E_INVALIDARG`
Always build players with `channels=2` and upmix. The synthesizer returns mono.

### Probe optional APIs in their own try
An absent function raising inside the main setup path took down the entire voice and fell
back to SAPI silently. A probe must never be able to break the path it is probing for.

## Rules

### Positioning: constant-power pan law, never linear balance
```python
theta = (pan + 50) / 100 * (math.pi / 2)
left, right = math.cos(theta), math.sin(theta)   # centre -> ~0.707
```
Our audio is mono duplicated into both channels, so it is fully correlated: linear
balance makes the centre up to 6 dB louder, and it reads as a volume bug.

### Apply channel volume only when it changes
Setting it before every utterance produced audible bursts - WASAPI channel volume changes
are not sample accurate. Steady-state typing must set it zero times.

### The two voice engines share one interface
`_winrt.WinRTVoice` is primary, `_voice.SapiVoice` is the fallback. Every method the
plugin calls must exist on both with the same signature. Letting them drift apart once
meant a `TypeError` on every keystroke and no character echo at all.

### Do not reimplement NVDA's echo logic
`_echo.py` runs NVDA's own `speakTypedCharacters` and swaps the two output functions.
NVDA keeps the typed-word buffer, protected fields, the suppression counter and terminal
handling - state that is not all visible from outside. Take only what it decided to say.

### Word echo belongs to NVDA
Only the character echo uses our voice. Word echo already works through NVDA's voice at
the user's own rate and rate boost; re-synthesising it caused more problems than it
solved.

### Ring announcements come from the stream being tuned
Adjusting the character voice while listening to the main voice tells you nothing about
what you are changing. The announcement *is* the sample - do not add a separate preview.

### Never leave a setting with a one-way door
The error volume's "follow NVDA" is reachable again below zero. A default you can leave
but not return to is a trap.

## Before deploying

Twice now a text-based edit has silently deleted a function or broken a `try` block, and
both were found only through the user's crash logs. Parse every module, collect what each
one defines, and verify every cross-module reference resolves - then syntax-check.

`_winrt.py` imports only `logHandler` and `._audio`, so it runs in plain Python with those
two stubbed. Test there before touching a live screen reader.

## Deployment

Tag format: `spatial-typing-feedback-vX.X.X-beta` (plugin prefix REQUIRED)

Minimum NVDA 2025.1: `speech.extensions.filter_speechSequence` does not exist in 2024.1,
and without it ring announcements come through twice.

## Current status

Working: three streams, the WinRT character voice, settings ring tuning, persistence.

Outstanding work is in `docs/TODOs/next-steps.md`. The most important item is that the
add-on does not yet respect NVDA's speech mode or sleep mode - it keeps talking after the
screen reader has been silenced.
