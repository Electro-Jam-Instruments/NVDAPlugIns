# Architecture Decisions - Windows Dictation Silence Plugin

## Decision 1: GlobalPlugin vs AppModule

**Choice:** GlobalPlugin

**Why:**
- Voice Typing works across ALL applications
- Need to intercept Win+H system-wide
- AppModule is app-specific, GlobalPlugin is system-wide

## Decision 2: Detection Method

### Options Considered

#### Option A: Focus-based detection (v0.0.1 - FAILED)
```python
def event_gainFocus(self, obj, nextHandler):
    if self._is_voice_typing_window(obj):
        # silence speech
```

**Why it failed:** Voice Typing is an overlay that doesn't take focus. Focus stays in the original text field. No focus event fires.

#### Option B: Timer-based polling (v0.0.2 - WORKS but undesirable)
```python
def script_toggleVoiceTyping(self, gesture):
    speech.setSpeechMode(speech.SpeechMode.off)
    gesture.send()
    core.callLater(500, self._poll_voice_typing_state)

def _poll_voice_typing_state(self):
    if find_voice_typing_window():
        core.callLater(300, self._poll_voice_typing_state)
    else:
        self._restore_speech()
```

**Why undesirable:** User doesn't like timers. Adds latency. Wastes CPU cycles.

#### Option C: Keypress interception (PREFERRED - TODO)
```python
def script_toggleVoiceTyping(self, gesture):
    speech.setSpeechMode(speech.SpeechMode.off)
    gesture.send()
    # Hook all keypresses
    self._hook_keyboard()

def _keyboard_hook(self, key):
    # Any key closes Voice Typing
    self._restore_speech()
    self._unhook_keyboard()
    # Pass key through
```

**Why preferred:**
- No timers
- Instant response
- Based on actual Voice Typing behavior (any key closes it)

**Choice:** Implement Option C

## Decision 3: How to Hook Keypresses in NVDA

### Options

#### A: inputCore.decide_executeGesture
NVDA's gesture handling system. Can intercept before gestures execute.

#### B: winInputHook
Low-level Windows keyboard hook via ctypes.

#### C: Custom keyboard filter
Register a filter with NVDA's input system.

**Research needed:** Which approach works best for intercepting ANY keypress temporarily?

## Decision 4: Speech Mode Management

**Choice:** Save previous mode, restore on close

```python
self._previous_speech_mode = speech.getSpeechMode()
speech.setSpeechMode(speech.SpeechMode.off)
# ... later ...
speech.setSpeechMode(self._previous_speech_mode)
```

**Why:** User might have been in "beeps" mode or other non-talk mode. We should restore exactly what they had.

## Decision 5: Window Detection Method

**Choice:** EnumWindows + process name check

```python
def find_voice_typing_window():
    # Enumerate all windows
    # Find Windows.UI.Core.CoreWindow class
    # Check if process is TextInputHost.exe
    return found
```

**Why:**
- Voice Typing window class: `Windows.UI.Core.CoreWindow`
- Process: `TextInputHost.exe`
- Can't use UIA focus because it doesn't take focus

## Open Questions

1. **Does Escape key close Voice Typing?** Need to test - may need special handling
2. **Does clicking the X button close it?** If so, we need mouse click detection too
3. **What about the microphone button?** Clicking it toggles listening state
4. **Edge case: What if user presses Win+H again?** Should toggle speech back on

## Technical Notes

### Voice Typing Process Info
- Process: `TextInputHost.exe`
- Window class: `Windows.UI.Core.CoreWindow`
- Does NOT take focus
- Overlay on top of other windows
- Auto-closes on any keyboard input

### NVDA Speech Modes
```python
speech.SpeechMode.off      # No speech
speech.SpeechMode.beeps    # Beeps instead of speech
speech.SpeechMode.talk     # Normal speech
```

### Key NVDA APIs
```python
import speech
speech.getSpeechMode()           # Get current mode
speech.setSpeechMode(mode)       # Set mode

from scriptHandler import script
@script(gesture="kb:windows+h")  # Bind to hotkey

import core
core.callLater(ms, func)         # Delayed call (timer)
```
