# User Requirements - Windows Dictation Silence Plugin

## Problem

When using Windows Voice Typing (Win+H), NVDA reads back the dictated text as it's inserted into the text field. The user hears everything they say repeated back, which is disruptive.

## User's Exact Words

> "I simply want to turn off word echo / speech when using WinKey + H -- my issue is that when using WinKey + H I am hearing back what is being put into the edit field -- yes I can turn off speech manually but I don't want to have to think about it."

> "yes any key press turns it off - so I think being confident when it started is the key here."

> "I'm not a fan of a timer here"

## Requirements

### Must Have

1. **Auto-silence on Win+H** - When user presses Win+H to start Voice Typing, NVDA speech automatically turns off
2. **Auto-restore on close** - When Voice Typing closes (any keypress), speech automatically restores
3. **No timers** - User explicitly doesn't want polling/timer-based solutions
4. **Transparent** - Should "just work" without user thinking about it

### Nice to Have

1. Restore exact previous speech mode (not just "talk")
2. Handle edge cases gracefully (Escape, clicking X, etc.)
3. Log useful debugging info

### Out of Scope (for now)

1. Suppressing specific UIA events (NVDA core would need to handle this)
2. Integration with Voice Access (different from Voice Typing)
3. Supporting older Windows versions

## User's Environment

- Windows 11 (latest)
- NVDA (latest public release)
- NVDA Settings:
  - "Registration for UI Automation events and property changes: Automatic (prefer selective)"
  - Scratchpad enabled for testing

## Voice Typing Behavior (from user testing)

1. Win+H opens the Voice Typing overlay
2. Shows "Listening..." when actively listening
3. Shows microphone icon when not listening
4. **Any keypress dismisses Voice Typing**
5. Focus stays in the original text field (overlay doesn't take focus)

## Success Criteria

1. User presses Win+H
2. Voice Typing opens
3. NVDA is silent (no text echo)
4. User dictates text
5. User presses any key (or text auto-completes)
6. Voice Typing closes
7. NVDA speech resumes automatically
8. User didn't have to manually toggle NVDA speech mode
