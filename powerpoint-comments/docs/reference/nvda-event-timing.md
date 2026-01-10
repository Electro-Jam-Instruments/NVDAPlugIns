# NVDA Event Timing Reference

Understanding the order of events is critical for modifying object properties before they're announced.

## Event Order (When Object Gains Focus)

| Order | Event/Method | Thread | Purpose |
|-------|--------------|--------|---------|
| 1 | `chooseNVDAObjectOverlayClasses(obj, clsList)` | Main | Select overlay classes |
| 2 | `event_NVDAObject_init(obj)` | Main | Modify object properties |
| 3 | `event_gainFocus(obj, nextHandler)` | Main | Handle focus, trigger announcements |

## event_NVDAObject_init - Modify Before Announcement

**Key insight:** Fires BEFORE NVDA announces the object. Modify `obj.name` here and NVDA speaks the modified name.

```python
def event_NVDAObject_init(self, obj):
    """Modify object properties BEFORE announcement."""
    window_class = getattr(obj, 'windowClassName', '')
    name = getattr(obj, 'name', '') or ''

    if window_class == 'mdiClass' and name.startswith('Slide '):
        obj.name = f"has notes, {name}"
```

## Timing Challenge with Worker Threads

**Problem:** Main thread events fire BEFORE worker thread has queried new data:

```
Timeline (slide navigation):
├─ 0ms:   COM WindowSelectionChange event fires (worker thread)
├─ 0ms:   event_NVDAObject_init fires (main thread) ← STALE DATA
├─ 0ms:   event_gainFocus fires (main thread) ← STALE DATA
├─ 23ms:  Worker thread finishes COM query, updates cache
```

**Solutions:**

1. **Lazy evaluation** - Use `_get_name()` in overlay class (queries when needed)
2. **COM events that fire earlier** - WindowSelectionChange fires before NVDA events
3. **Synchronous query** - Query COM in `event_NVDAObject_init` (may slow NVDA)

## Our Solution: Lazy _get_name()

```python
class CustomSlide(NVDAObject):
    def _get_name(self):
        """Called lazily when NVDA actually needs the name."""
        # By now, worker thread has finished querying
        prefix = worker._cached_prefix or ""
        original_name = self._original_name
        if prefix:
            return f"{prefix}, {original_name}"
        return original_name
```

## Focus Event Sequence (Multiple Objects)

```
Old Focus → New Focus:
1. loseFocus on old focus
2. focusExited on old focus's parent (up to common ancestor)
3. focusEntered on new ancestors (down from common ancestor)
4. gainFocus on new focus
```

## COM Event Timing (Our Advantage)

COM events fire as PowerPoint processes changes, BEFORE NVDA detects them via UIA:

```
User presses Page Down:
├─ T+0ms:   PowerPoint changes slide internally
├─ T+1ms:   COM WindowSelectionChange fires (our sink receives it)
├─ T+5ms:   Worker thread queries and caches slide data
├─ T+15ms:  NVDA detects UIA event
├─ T+15ms:  event_NVDAObject_init fires
├─ T+15ms:  _get_name() called → reads FRESH cached data
├─ T+20ms:  User hears correct announcement
```

## Use Cases Summary

| Event | Use When |
|-------|----------|
| `chooseNVDAObjectOverlayClasses` | Apply custom class with multiple overrides |
| `event_NVDAObject_init` | Simple property changes (name, role) |
| `event_gainFocus` | React to focus, trigger side effects |
| Overlay `_get_name()` | Need lazy/fresh data at announcement time |

## Preventing Premature Announcements - _has_received_focus Flag

**Problem:** When NVDA starts with PowerPoint already open but not focused, the addon would announce slide info even though user is in another app.

**Solution:** Track whether PowerPoint has actually received focus:

```python
class PowerPointWorker:
    def __init__(self):
        self._has_received_focus = False  # Track real focus

    def request_initialize(self):
        """Called from event_appModule_gainFocus."""
        self._has_received_focus = True  # Now we know app is focused

    def _check_initial_slide(self):
        """Called during initialization."""
        # Don't announce if app hasn't received focus yet
        if not self._has_received_focus:
            log.info("Skipping announcement - app not focused yet")
            return  # DON'T mark as announced either!
```

This ensures announcements only happen when user actually switches to PowerPoint.

## Speech Cancellation - _just_navigated Flag

**Problem:** When reformatting comments, we cancel NVDA's default speech and announce our version. But after slide navigation (PageUp/PageDown), this would cut off the slide title announcement.

**Solution:** Skip speech cancellation after navigation:

```python
class AppModule(AppModule):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._just_navigated = False

    def script_navigateSlide(self, gesture):
        """Handle PageUp/PageDown."""
        self._just_navigated = True  # Set flag before navigation
        # ... navigate ...

    def event_gainFocus(self, obj, nextHandler):
        # When reformatting comments...
        if reformatted_name:
            # DON'T cancel speech if we just navigated (preserves slide title)
            if not self._just_navigated:
                speech.cancelSpeech()
            else:
                self._just_navigated = False  # Reset flag

            ui.message(reformatted_name)
            return

        nextHandler()
```

**Timeline with flag:**
```
User presses PageDown in Comments pane:
├─ _just_navigated = True
├─ Worker navigates slide
├─ Slide title announced by NVDA
├─ Focus returns to comment
├─ event_gainFocus fires
├─ Check _just_navigated → True → skip cancelSpeech()
├─ _just_navigated = False
├─ Announce reformatted comment
└─ User hears: "Slide 2 Title" ... "Author:"
```

Without the flag, user would only hear "Author:" (slide title cut off).
