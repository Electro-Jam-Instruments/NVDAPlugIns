# Changelog - Windows Dictation Silence Plugin

All notable changes to the Windows Dictation Silence plugin will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.0-beta] - 2026-01-11

Initial public beta release. Full stable release coming soon.

### Added
- **Auto-silence on Win+H** - NVDA goes silent when Windows Voice Typing opens
- **Instant restore** - Speech returns immediately on any keypress
- **Mode preservation** - Restores your previous speech mode (talk, beeps, or off)

### Technical
- GlobalPlugin architecture for system-wide operation
- Gesture filter approach using `inputCore.decide_executeGesture` (no polling/timers)
- Standard NVDA scons build system

---

## Version History (Development)

Pre-release development versions (internal testing):

| Version | Notes |
|---------|-------|
| 0.0.7 | Migrated to standard scons build system |
| 0.0.6 | Fixed speech API - use getState().speechMode |
| 0.0.3 | Keypress interception approach (no timers) |
| 0.0.2 | Timer-based polling (deprecated - user requested no timers) |
| 0.0.1 | Focus-based detection (failed - Voice Typing doesn't take focus) |
