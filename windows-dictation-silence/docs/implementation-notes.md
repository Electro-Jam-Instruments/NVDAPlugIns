# Implementation Notes - Windows Dictation Silence Plugin

## Current Implementation: v0.0.3 (Keypress Interception)

### How It Works

1. **Win+H Interception** (lines 51-67):
   - Plugin binds `@script(gesture="kb:windows+h")`
   - When pressed: calls `_start_voice_typing_mode()`
   - Passes Win+H through via `gesture.send()`

2. **Speech Silencing** (lines 69-75):
   - Saves current speech mode via `speech.getSpeechMode()`
   - Sets speech to OFF via `speech.setSpeechMode(speech.SpeechMode.off)`
   - Installs gesture filter

3. **Gesture Filter** (lines 111-133):
   - Registered with `inputCore.decide_executeGesture.register()`
   - Called for EVERY gesture (keyboard, mouse, etc.)
   - On any key (except Win+H): restores speech, removes filter
   - Returns True to allow gesture through

4. **Speech Restoration** (lines 77-85):
   - Removes gesture filter
   - Restores saved speech mode

### Code Structure (v0.0.3)

```
windowsDictationSilence.py
└── GlobalPlugin
    ├── __init__()                   # Initialize state
    ├── terminate()                  # Cleanup on unload
    ├── script_toggleVoiceTyping()   # Win+H handler
    ├── _start_voice_typing_mode()   # Silence speech + install filter
    ├── _end_voice_typing_mode()     # Restore speech + remove filter
    ├── _install_gesture_filter()    # Hook into inputCore
    ├── _remove_gesture_filter()     # Unhook from inputCore
    └── _gesture_filter()            # Filter callback - any key restores speech
```

### State Variables (v0.0.3)

```python
self._previous_speech_mode = None      # Saved mode to restore
self._voice_typing_active = False      # Currently in Voice Typing?
self._gesture_filter_installed = False # Filter currently active?
```

### Why This Approach

- **No timers** - User explicitly requested no polling
- **Instant response** - Speech restores immediately on keypress
- **Based on actual behavior** - Voice Typing closes on any key
- **Clean lifecycle** - Filter is installed/removed as needed

---

## Previous Implementation: v0.0.2 (Timer-Based) - DEPRECATED

### Why It Was Replaced

1. **Timer-based polling** - User explicitly doesn't want timers
2. **300ms latency** - Delay between Voice Typing closing and speech restoring
3. **CPU overhead** - Constantly enumerating windows

### How It Worked

- Polled every 300ms using `core.callLater()`
- Enumerated windows looking for TextInputHost.exe
- Restored speech when window disappeared

---

## Edge Cases to Handle

## Testing Checklist

- [ ] Win+H opens Voice Typing and silences NVDA
- [ ] Typing a letter closes Voice Typing and restores speech
- [ ] Pressing Escape closes Voice Typing and restores speech
- [ ] Win+H again closes Voice Typing and restores speech
- [ ] Speech mode restored to exact previous state (not just "talk")
- [ ] No errors in NVDA log
- [ ] Plugin loads cleanly on NVDA restart
- [ ] Plugin unloads cleanly

## Research Needed

1. How does `inputCore.decide_executeGesture` work exactly?
2. Can we register/unregister handlers dynamically?
3. Are there NVDA addons that do similar keypress interception?
4. Test: Does clicking the microphone button count as "any key"?

## Log Messages

Plugin logs to NVDA log. Check with NVDA+F1:

```
INFO - globalPlugins.windowsDictationSilence: Windows Dictation Silence: Plugin initialized
INFO - globalPlugins.windowsDictationSilence: Windows Dictation Silence: Win+H pressed
INFO - globalPlugins.windowsDictationSilence: Windows Dictation Silence: Speech OFF
INFO - globalPlugins.windowsDictationSilence: Windows Dictation Silence: Speech restored
```
