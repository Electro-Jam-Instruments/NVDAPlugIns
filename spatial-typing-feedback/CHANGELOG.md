# Changelog - Spatial Typing Feedback Plugin

All notable changes to the Spatial Typing Feedback plugin will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.0-beta] - 2026-08-30

First beta. Three streams, positioned and levelled.

### Added
- **Characters off to the right, quieter** - each typed character is spoken by a second
  Windows OneCore voice at +35 right and 50% volume, so the echo stays out of the way
  while your main voice keeps reading
- **Completed words stay centre** - the word echo is spoken by that second voice at
  centre and full volume, sharing the position with NVDA's main speech
- **Typing errors on the left** - NVDA's spelling alert plays at -35 left and 50% volume,
  mirroring the character echo on the other side
- **The second voice matches your main one** - same OneCore voice, rate, pitch and
  **rate boost**, so the echo never lags behind a boosted main voice. Streams are told
  apart by position, not by sounding different
- `NVDA+shift+p` - play a test phrase at each stream position
- `NVDA+shift+s` - toggle the whole thing on and off

### Technical
- GlobalPlugin architecture for system-wide operation
- Secondary voice is a second, independent `ocSpeech` token rendered into a buffer we
  own, so panning, level and lifetime are ours. OneCore is token-based, so instances do
  not interfere with NVDA's own synthesizer
- Mono OneCore output is upmixed to stereo before positioning, because per-channel
  volume is meaningless on a mono player
- Constant-power pan law, since a mono signal duplicated into both channels is up to
  6 dB louder at centre than panned to one side
- Character and word echo are intercepted at `speech.speakTypedCharacters`, where NVDA
  decides both. NVDA's own logic runs untouched - typed-word buffer, protected fields,
  suppression counter, terminal handling - and only its output is captured
- Error alert is redirected through the `nvwave.decide_playWaveFile` extension point,
  and only for the typing-time alert. Reading-time error reports stay with the main voice
- Goes dormant if NVDA's Sound Split is enabled and says so; the two features are
  mutually exclusive

### Known limits
- Requires Windows OneCore voices and a stereo output device
- Stream positions and levels are fixed in this build. A settings panel and settings
  ring slots come once the values have been judged by ear
- The main voice itself is not yet movable; it sits at centre
