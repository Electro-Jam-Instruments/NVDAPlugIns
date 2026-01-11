# Windows Dictation Silence Plugin - Development Instructions

## Critical Patterns

These patterns are MANDATORY for this plugin.

### GlobalPlugin (not AppModule)
```python
from globalPluginHandler import GlobalPlugin

class GlobalPlugin(GlobalPlugin):
    pass
```
**Why:** Voice Typing works across ALL apps. AppModule is app-specific, GlobalPlugin is system-wide.

See: `docs/architecture-decisions.md` Decision #1

### Speech Mode Management
```python
# Save before changing
self._previous_speech_mode = speech.getSpeechMode()
speech.setSpeechMode(speech.SpeechMode.off)

# Restore exactly what user had
speech.setSpeechMode(self._previous_speech_mode)
```
**Why:** User might be in "beeps" mode. Always restore exact previous state.

See: `docs/architecture-decisions.md` Decision #4

### Gesture Filter for Keypress Detection
```python
inputCore.decide_executeGesture.register(self._gesture_filter)
# ... later ...
inputCore.decide_executeGesture.unregister(self._gesture_filter)
```
**Why:** No timers/polling. Any keypress closes Voice Typing, so intercept and restore speech immediately.

See: `docs/architecture-decisions.md` Decision #2 (Option C), `docs/implementation-notes.md`

## Documentation

| Need | Location |
|------|----------|
| Why decisions were made | `docs/architecture-decisions.md` |
| Current implementation details | `docs/implementation-notes.md` |
| User requirements | `docs/user-requirements.md` |
| Deployment plan | `docs/deployment-plan.md` |

## Key Technical Facts

- **Target:** Windows Voice Typing (Win+H)
- **Process:** `TextInputHost.exe`
- **Window class:** `Windows.UI.Core.CoreWindow`
- **Behavior:** Overlay, doesn't take focus, closes on any keypress

## Deployment

Tag format: `windows-dictation-silence-vX.X.X-beta` (plugin prefix REQUIRED)

Use `/deploy` command or see `../.claude/commands/deploy.md`

## Current Status

v0.0.3 - Keypress interception approach (no timers)
