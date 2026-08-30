# Spatial Typing Feedback - Developer Documentation

This folder contains technical documentation for the Spatial Typing Feedback NVDA addon.

## Documentation Structure

| File | Purpose |
|------|---------|
| `user-requirements.md` | Problem statement and the requirements (R1-R9) |
| `architecture-decisions.md` | Eleven decisions (D1-D11) with rationale, plus the phase plan |
| `research/nvda-audio-and-typing-hooks.md` | Verified NVDA source evidence behind every decision |
| `research/spatial-audio-options.md` | Future option only: what 3D HRTF would cost and deliver |
| `TODOs/` | Open work items |

## Key Documents

### Requirements
- **[user-requirements.md](user-requirements.md)** - R1 character echo right, R2 errors left, R3 completed words centre, R4 clean off switch, R5 configurable, R6 secondary voice matches the main voice (incl. rate boost), R7 settings ring slots for every stream, R8 multi-stream master on/off

### Architecture
- **[architecture-decisions.md](architecture-decisions.md)** - D1 GlobalPlugin, D2 panning via `WavePlayer.setVolume`, D3 secondary voice via a second OneCore token (not SAPI 5 - corrected), D4 wrapping `speakTypedCharacters` (covers **both** R1 and R3), D5 `decide_playWaveFile` for the error alert, D6 autocomplete ruled out of scope, D7 Sound Split is mutually exclusive, D8 settings ring integration and pan law, D9 panning the main voice, D10 pan and volume only (3D deferred), D11 no settings layer until the values are heard

### Research
- **[research/nvda-audio-and-typing-hooks.md](research/nvda-audio-and-typing-hooks.md)** - Quoted NVDA source for every hook, with file and line references, plus the open questions spike S1 must answer

## When to Use Each Section

| If you need to... | Look at... |
|-------------------|------------|
| Understand what problem we're solving | `user-requirements.md` |
| Understand why an approach was chosen or rejected | `architecture-decisions.md` |
| Find the exact NVDA API a decision depends on | `research/nvda-audio-and-typing-hooks.md` |
| Know what to build next | The phase table at the end of `architecture-decisions.md` |
| Understand why autocomplete is *not* in scope | Decision D6 |

## Highest-Risk Areas

Read these before writing code:

1. **D4 is a monkeypatch, and it is the whole of R1 and R3.** Wrapping `speech.speakTypedCharacters` is version-fragile. It must fail safe - if the patch cannot be applied, NVDA must behave exactly as it did before.
2. **D3 rests on an unproven assumption.** A second `ocSpeech_initialize` in one process *should* work - the interface is token-based - but it has not been tested. Spike S1 proves it before anything is built on it.
3. **OneCore returns mono; panning needs stereo.** `setVolume(left=, right=)` raises `E_INVALIDARG` on channel 1 of a mono player. The upmix is mandatory and easy to forget.
4. **Character echo and word echo can fire on the same keystroke.** Pressing space completes a word *and* echoes as a character. The secondary voice queue must handle two overlapping utterances at two different pan positions.
5. **NVDA 2026.1 renamed `WasapiWavePlayer` to `WavePlayer`** and changed the `outputDevice` parameter. Version guards must handle both.
