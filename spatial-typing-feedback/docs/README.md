# Spatial Typing Feedback - developer documentation

Describes what exists, not what was once planned.

| File | Purpose |
|------|---------|
| [`architecture.md`](architecture.md) | How it works and why. Read this first. |
| [`research/nvda-and-winrt-reference.md`](research/nvda-and-winrt-reference.md) | The exact NVDA and Windows surfaces it depends on: IIDs, vtable slots, measurements |
| [`TODOs/next-steps.md`](TODOs/next-steps.md) | What is left |

## Code layout

| File | Holds |
|------|-------|
| `__init__.py` | GlobalPlugin: lifecycle, hooks, stream settings, ring proxies |
| `_streams.py` | The three streams and their starting values |
| `_config.py` | Persisted settings; saved on leaving the ring, not on a timer |
| `_audio.py` | `PannedPlayer`: pan law, mono-to-stereo upmix, channel gains |
| `_winrt.py` | The character voice - a Windows SpeechSynthesizer we activate ourselves |
| `_voice.py` | SAPI 5 fallback, same interface |
| `_echo.py` | Wraps `speech.speakTypedCharacters` |
| `_errors.py` | `decide_playWaveFile` hook for the typing-error alert |
| `_ring.py` | Settings ring slots |

## Before changing anything

Four things here are not what they appear, and each cost hours to find. Explained in
`architecture.md`; in short:

1. **`SpeechSynthesizerOptions.SpeakingRate` does nothing.** It stores and reads back
   correctly. Rate is set through SSML prosody instead.
2. **SSML prosody rate saturates at 200%.** Larger numbers silently pin the voice at
   maximum, which sounds like a broken mapping.
3. **Probe optional APIs in their own try.** An absent function raising inside the main
   setup path took down the whole voice and fell back silently to SAPI.
4. **`setVolume` on a mono player raises `E_INVALIDARG`.** Always upmix to stereo first.

## Testing without NVDA

`_winrt.py` imports only `logHandler` and `._audio`. Stub those two and it runs in plain
Python - which is how voices, rate, pitch, voice selection, interruption and shutdown were
all verified before going near a live screen reader. Worth keeping it that way.
