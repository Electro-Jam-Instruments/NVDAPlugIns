# Changelog - Spatial Typing Feedback Plugin

All notable changes to the Spatial Typing Feedback plugin will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

Nothing tagged yet. Everything below is in the working tree.

### Fixed - 0.1.1
- **NVDA no longer crashes mid-session.** Three memory-safety faults on the speech
  path, all the same mistake: treating an asynchronous Windows call as finished when
  it returned. The SSML string was freed while the engine was still reading it; the
  `_await` timeout abandoned operations still in flight and the caller then released
  the buffer being written into; and a successful voice change leaked the entire voice
  collection. Symptom was `nvda.exe` dying inside `MSTTSEngine_OneCore.dll` with
  `0xc0000409` after hours of use, with nothing in NVDA's log.
- **Respects NVDA's speech mode and sleep mode.** Silencing NVDA left the character
  echo talking, which made a quiet main voice look like a broken add-on.
- A stalled speech operation is now reported in the normal log rather than only under
  debug logging.

### Added
- **Characters off to the right, quieter** - each typed character spoken by a second
  voice, positioned and levelled separately from NVDA's own speech
- **Typing errors on the left** - NVDA's spelling alert, repositioned
- **Spelling notes on the left** - NVDA's spoken "spelling error" / "out of spelling
  error" moved to their own voice and position, so they inform rather than interrupt
- **A second voice of its own** - independent voice, speed, pitch, rate boost and
  punctuation pauses, drawn from your installed Windows voices
- **Tuning from NVDA's settings ring** - pick a stream, adjust it, and hear each change
  immediately in the voice being tuned
- Settings persist, saved when you leave the ring
- `NVDA+shift+p` plays a test at each position; `NVDA+shift+s` toggles the add-on

### Technical
- The character voice activates `Windows.Media.SpeechSynthesis.SpeechSynthesizer`
  directly, so it offers the same voices as NVDA without depending on NVDA's internals
- Constant-power pan law over per-channel volume, with mono upmixed to stereo first
- Typed characters and completed words are separated at `speech.speakTypedCharacters`,
  leaving NVDA's own word buffer and suppression logic untouched
- The error alert is redirected through the `nvwave.decide_playWaveFile` extension point
- Goes dormant when NVDA's Sound Split is enabled; the two are mutually exclusive

### Known limitations
- Replaces NVDA's own settings ring slots rather than adding to them
- Requires Windows voices and a stereo output device
