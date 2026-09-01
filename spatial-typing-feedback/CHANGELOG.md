# Changelog - Spatial Typing Feedback Plugin

All notable changes to the Spatial Typing Feedback plugin will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

Nothing released yet. Everything below is in the working tree.

### Added
- **Characters off to the right, quieter** - each typed character spoken by a second
  voice, positioned and levelled separately from NVDA's own speech
- **Typing errors on the left** - NVDA's spelling alert, repositioned
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
- Does not yet respect NVDA's speech mode or sleep mode
- Replaces NVDA's own settings ring slots rather than adding to them
- Requires Windows voices and a stereo output device
