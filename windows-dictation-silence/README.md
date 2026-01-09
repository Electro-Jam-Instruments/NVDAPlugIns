# Windows Dictation Silence - NVDA Plugin

Auto-silence NVDA speech when Windows Voice Typing (Win+H) is active.

## Problem Statement

When using Windows Voice Typing (Win+H), NVDA echoes back the dictated text as it appears in the text field. This creates an annoying feedback loop where you hear everything you say repeated back to you.

**User's goal:** When Win+H is pressed to start dictation, NVDA should automatically go silent. When dictation ends, speech should restore.

## Current Status: v0.0.2 - Work In Progress

The current implementation uses timer-based polling which the user doesn't like. Need to switch to keypress interception approach.

## Key Technical Findings

### Voice Typing Behavior (Windows 11)

1. **Win+H opens Voice Typing** - A floating overlay appears
2. **"Listening..." tooltip** shows when actively listening
3. **Any keypress closes it** - Voice Typing dismisses on any keyboard input
4. **Focus stays in original app** - The overlay doesn't take focus
5. **Text echo is the problem** - NVDA reads UIA name change events as text is inserted

### Why Built-in NVDA Fix Doesn't Work

NVDA issue #12938 mentions "selective UIA event registration" should suppress dictation text in Windows 11 22H2+. The user has:
- Latest public NVDA release
- Windows 11 latest
- Setting: "Registration for UI Automation events and property changes: Automatic (prefer selective)"

Yet text echo still occurs. This plugin is needed to fill the gap.

### Detection Challenges

1. **No focus event** - Voice Typing overlay doesn't take focus
2. **No close notification** - Windows doesn't broadcast when Voice Typing closes
3. **TextInputHost.exe** - Voice Typing UI runs in this process with `Windows.UI.Core.CoreWindow` class

## Proposed Approaches

### Approach A: Timer-Based Polling (Current - v0.0.2)
- Win+H → speech OFF → poll every 300ms for window existence
- **Pro:** Detects window close reliably
- **Con:** User doesn't like timers

### Approach B: Keypress Interception (Preferred)
- Win+H → speech OFF + hook keyboard
- Any subsequent keypress → restore speech + unhook + pass key through
- **Pro:** No timers, instant response
- **Con:** More complex, need to handle edge cases

### Approach C: Hybrid
- Win+H → speech OFF
- Monitor for specific close signals (Escape, clicking X, etc.)

## Files

```
windows-dictation-silence/
├── addon/
│   ├── manifest.ini              # Plugin metadata
│   └── globalPlugins/
│       └── windowsDictationSilence.py  # Main plugin code
├── docs/
│   ├── architecture-decisions.md # Technical decisions
│   └── implementation-notes.md   # Implementation details
└── README.md                     # This file
```

## Installation (Testing)

1. Copy `windowsDictationSilence.py` to `%APPDATA%\nvda\scratchpad\globalPlugins\`
2. Enable scratchpad in NVDA: Settings > Advanced > Enable loading custom code
3. Restart NVDA or press NVDA+Ctrl+F3
4. Test with Win+H

## Next Steps

1. Implement Approach B (keypress interception)
2. Test with real dictation workflow
3. Handle edge cases (Escape to close, click X, etc.)
4. Package as .nvda-addon

## References

- [NVDA Issue #12938: Silence NVDA while Dictation is actively listening](https://github.com/nvaccess/nvda/issues/12938)
- [DictationBridge Add-on](https://coolblindtech.com/dictation-bridge-brings-nvda-users-access-to-speech-recognition-software/)
- NVDA Global Plugin development guide
