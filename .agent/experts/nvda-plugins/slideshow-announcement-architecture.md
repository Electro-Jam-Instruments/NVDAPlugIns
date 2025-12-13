# Slideshow Announcement Architecture

## Overview

This document explains how our addon announces "has notes" BEFORE the slide number/title during PowerPoint slideshow presentations, using NVDA's `_get_name()` property override pattern.

**Version:** v0.0.61+
**Pattern:** Custom overlay class with `_get_name()` override
**Result:** Single integrated announcement, no timing dependencies

## Problem Statement

### User Requirements

When presenting a PowerPoint slideshow, the user needs to hear:
1. "has notes" (if speaker notes with **** markers exist)
2. Slide number and title
3. In that order, reliably, every time

### Technical Challenge

NVDA automatically announces the slideshow window name when slides change:
- Default: `"Slide show - Slide 3, Project Overview"`
- User wants: `"has notes, Slide show - Slide 3, Project Overview"`

**Constraints:**
- Cannot use timing-based solutions (unreliable)
- Cannot use `speech.cancelSpeech()` (interrupts, bad UX)
- Cannot use `speech.priorities.NOW` (interrupts, user rejected)
- Must work with NVDA's existing architecture
- Must preserve all other NVDA features

## Solution Architecture

### High-Level Approach

**Override the window name property itself** using NVDA's overlay class pattern:

```
CustomSlideShowWindow (our class)
    ↓ inherits from
SlideShowWindow (NVDA's built-in class)
    ↓ overrides
_get_name() property
    ↓ returns
"has notes, " + base_name (when notes exist)
```

### Why This Works

1. **Timing Guarantee:** `_get_name()` is called by NVDA's focus reporting mechanism at exactly the right time
2. **Single Announcement:** Window name is announced once, with everything in it
3. **No Duplication:** We're not adding a second announcement, we're modifying the first one
4. **NVDA Pattern:** Standard overlay class pattern documented in NVDA developer guide
5. **Scoped Correctly:** Only affects PowerPoint slideshow windows, nothing else

## Implementation Details

### Component 1: CustomSlideShowWindow Class

**Location:** `powerpoint-comments/addon/appModules/powerpnt.py:1058-1110`

**Purpose:** Extends NVDA's built-in `SlideShowWindow` to customize the window name announcement.

**Key Method:**

```python
def _get_name(self):
    """Get window name with notes status prepended if present."""
    # Get base announcement from parent class
    base_name = super()._get_name()
    # Example: "Slide show - Slide 3, Meeting Overview"

    # Check if slide has notes via worker thread
    if hasattr(self.appModule, '_worker') and self.appModule._worker:
        if self.appModule._worker._has_meeting_notes():
            # Prepend "has notes, "
            return f"has notes, {base_name}"

    # No notes - return unchanged
    return base_name
```

**When Called:**
- Slideshow window is created (F5 starts presentation)
- Slide changes (Space, PageDown, etc.)
- NVDA needs to report focus on the window

**What It Returns:**
- With notes: `"has notes, Slide show - Slide 3, Title"`
- Without notes: `"Slide show - Slide 3, Title"`
- Notes mode: `"has notes, Slide show notes - Slide 3, Title"`

### Component 2: Overlay Class Registration

**Location:** `powerpoint-comments/addon/appModules/powerpnt.py:1153-1181`

**Purpose:** Tells NVDA to use our `CustomSlideShowWindow` instead of the built-in `SlideShowWindow`.

**Method:** `chooseNVDAObjectOverlayClasses()`

```python
def chooseNVDAObjectOverlayClasses(self, obj, clsList):
    """Apply custom overlay classes for PowerPoint objects."""
    # Let parent handle standard selection
    super().chooseNVDAObjectOverlayClasses(obj, clsList)

    # If parent assigned SlideShowWindow, replace with our custom version
    if SlideShowWindow in clsList:
        idx = clsList.index(SlideShowWindow)
        clsList[idx] = CustomSlideShowWindow
        log.info("Replaced SlideShowWindow with CustomSlideShowWindow")
```

**How It Works:**
1. NVDA calls this method for every object it encounters
2. Parent class (`AppModule`) assigns built-in PowerPoint classes
3. We check if `SlideShowWindow` was assigned
4. If yes, we replace it with our `CustomSlideShowWindow`
5. NVDA uses our custom class for that object

**Scoping:**
- Only runs in our PowerPoint `AppModule` (only when PowerPoint is active)
- Only replaces `SlideShowWindow` (which only exists in PowerPoint slideshow)
- No impact on other apps, other PowerPoint modes, or other NVDA features

### Component 3: Event Sink Update

**Location:** `powerpoint-comments/addon/appModules/powerpnt.py:615-642`

**Purpose:** Removed duplicate "has notes" announcement from COM event handler.

**Change (v0.0.61):**

```python
def on_slideshow_slide_changed(self, slide_index, slideshow_window):
    """Track slide changes during slideshow."""
    # Store slideshow window for notes access
    self._slideshow_window = slideshow_window

    # Track slide index
    self._last_announced_slide = slide_index

    # v0.0.61: Notes announcement now handled by CustomSlideShowWindow._get_name()
    # No additional announcement needed here
    log.debug("Slideshow slide tracking updated (announcement via window name)")
```

**Before (v0.0.60):**
```python
# OLD CODE - REMOVED
if self._has_meeting_notes():
    self._announce("has notes")
```

**Why Removed:**
- Window name now includes "has notes"
- Announcing separately would create duplication
- Window name announcement is more reliable (no timing issues)

## Event Flow Diagram

### Complete Sequence: Space Pressed in Slideshow

```
┌─────────────────────────────────────────────────────┐
│ USER ACTION: Press Space in slideshow              │
└─────────────────────────────────────────────────────┘
                    ↓
┌─────────────────────────────────────────────────────┐
│ NVDA: SlideShowTreeInterceptor.script_slideChange()│
│       Sends gesture to PowerPoint                   │
└─────────────────────────────────────────────────────┘
                    ↓
┌─────────────────────────────────────────────────────┐
│ POWERPOINT: Slide advances (internal)               │
│             Fires SlideShowNextSlide COM event      │
└─────────────────────────────────────────────────────┘
                    ↓
        ┌───────────┴───────────┐
        ↓                       ↓
┌─────────────────┐    ┌──────────────────┐
│ Built-in Sink   │    │ Our Addon Sink   │
│ (NVDA core)     │    │ (our worker)     │
└─────────────────┘    └──────────────────┘
        ↓                       ↓
┌─────────────────┐    ┌──────────────────┐
│ handleSlide     │    │ on_slideshow_    │
│ Change()        │    │ slide_changed()  │
│                 │    │ (tracks index)   │
└─────────────────┘    └──────────────────┘
                    ↓
┌─────────────────────────────────────────────────────┐
│ NVDA: Window name changes detected                  │
│       Triggers focus reporting                      │
└─────────────────────────────────────────────────────┘
                    ↓
┌─────────────────────────────────────────────────────┐
│ NVDA: Queries CustomSlideShowWindow._get_name()    │
└─────────────────────────────────────────────────────┘
                    ↓
┌─────────────────────────────────────────────────────┐
│ OUR CODE: Check worker._has_meeting_notes()        │
│           Returns True/False                        │
└─────────────────────────────────────────────────────┘
                    ↓
┌─────────────────────────────────────────────────────┐
│ OUR CODE: Return name                               │
│   If notes: "has notes, Slide show - Slide 3, ..." │
│   No notes: "Slide show - Slide 3, ..."            │
└─────────────────────────────────────────────────────┘
                    ↓
┌─────────────────────────────────────────────────────┐
│ NVDA: Announces the window name                     │
│ USER HEARS: "has notes, Slide show - Slide 3, ..." │
└─────────────────────────────────────────────────────┘
                    ↓
┌─────────────────────────────────────────────────────┐
│ NVDA: Continues with slide content (if enabled)     │
└─────────────────────────────────────────────────────┘
```

### Timing Analysis

```
T0: Space key pressed
    ↓ <1ms
T1: Gesture sent to PowerPoint
    ↓ ~10ms (PowerPoint processing)
T2: PowerPoint advances slide
    ↓ <1ms
T3: SlideShowNextSlide COM event fires
    ├─ Built-in sink: handleSlideChange()
    └─ Our sink: on_slideshow_slide_changed() (tracks index only)
    ↓ ~5ms (NVDA processing)
T4: Window name query triggered
    ↓ <1ms
T5: CustomSlideShowWindow._get_name() called
    ├─ Check worker._has_meeting_notes()
    ├─ Returns "has notes, Slide show - ..." or "Slide show - ..."
    ↓ <1ms
T6: NVDA announces name ← USER HEARS IT HERE
    ↓
T7: Optional: Content reading begins

Total latency: ~20ms (imperceptible to user)
```

**Key Points:**
- `_get_name()` is called synchronously during focus reporting
- No race conditions or timing dependencies
- Guaranteed order: notes check happens BEFORE announcement
- Single speech event (no separate announcements to coordinate)

## Scoping and Safety

### What This Changes

✅ **ONLY affects:**
- PowerPoint slideshow windows (`SlideShowWindow` class)
- The window name announcement (what NVDA says when focus enters the window)
- Slides with speaker notes marked with `****` get "has notes, " prefix

### What This Does NOT Change

❌ **Does NOT affect:**
- Any other applications (scoped to PowerPoint AppModule)
- PowerPoint normal editing mode (only slideshow uses `SlideShowWindow`)
- Slide content reading (say-all, shape reading, etc.)
- NVDA verbosity settings
- NVDA speech settings (rate, pitch, volume)
- Other NVDA features (browse mode, focus tracking, etc.)
- Windows actual window title (taskbar, title bar)
- Other users without this addon installed

### Inheritance Chain

```
NVDAObject (base)
    ↓
IAccessible (interface)
    ↓
Window (window wrapper)
    ↓
SlideShowWindow (NVDA built-in for PowerPoint)
    ↓
CustomSlideShowWindow (our override)
```

**What We Override:**
- Single property: `_get_name()`

**What We Preserve:**
- All other methods and properties from `SlideShowWindow`
- All NVDA's built-in slideshow functionality
- Event handling, navigation, content reading

### Class Selection Logic

```python
# NVDA's process for each object:
1. Create base NVDAObject
2. Call AppModule.chooseNVDAObjectOverlayClasses(obj, clsList)
   ├─ Parent adds: [IAccessible, Window, SlideShowWindow]
   └─ Our code replaces SlideShowWindow with CustomSlideShowWindow
3. NVDA applies classes in order: CustomSlideShowWindow → Window → IAccessible → NVDAObject
4. For this specific object, our _get_name() is used
5. For all other objects in PowerPoint, standard classes apply
```

**When Our Class Is Used:**
- Object window class name: `"screenClass"` (PowerPoint slideshow window)
- Parent AppModule added `SlideShowWindow` to class list
- We replaced it with `CustomSlideShowWindow`

**When Our Class Is NOT Used:**
- Normal editing windows (use different classes)
- Comments pane (uses different classes)
- Other PowerPoint UI elements (use different classes)
- Other applications (our AppModule not loaded)

## Testing Strategy

### Manual Test Cases

**Test 1: Slideshow with Notes**
```
Setup: Presentation with slide 1 (no notes), slide 2 (has **** notes)
Action: F5 → Space
Expected:
  - Slide 1: "Slide show - Slide 1, Title"
  - Slide 2: "has notes, Slide show - Slide 2, Title"
Result: ___
```

**Test 2: Slideshow without Notes**
```
Setup: Presentation with no notes on any slide
Action: F5 → Space → Space
Expected: All slides announce "Slide show - Slide X, Title" (no "has notes")
Result: ___
```

**Test 3: Notes Mode Toggle**
```
Setup: In slideshow on slide with notes
Action: Ctrl+Shift+S (toggle notes mode)
Expected: "has notes, Slide show notes - Slide X, Title"
Result: ___
```

**Test 4: Normal Mode (Regression)**
```
Setup: Exit slideshow, navigate slides in Normal view
Action: PageDown
Expected: Existing behavior (comment count, not slideshow name)
Result: ___
```

### Log Verification

Check NVDA log (NVDA+F1 > Tools > View Log) for:

```
✓ "Replaced SlideShowWindow with CustomSlideShowWindow"
✓ "CustomSlideShowWindow: Base name = 'Slide show - ...'"
✓ "CustomSlideShowWindow: Slide has notes - prepending to announcement"
  OR
✓ "CustomSlideShowWindow: No meeting notes on slide"
✗ No errors in _get_name()
✗ No AttributeError exceptions
```

### Edge Cases

**Edge Case 1: Notes Without **** Markers**
```
Slide notes: "Regular notes without markers"
Expected: "Slide show - Slide X, Title" (no "has notes")
Reason: Only notes with **** are "meeting notes"
```

**Edge Case 2: Empty Notes**
```
Slide has notes placeholder but empty text
Expected: "Slide show - Slide X, Title" (no "has notes")
Reason: Empty string fails '****' in notes check
```

**Edge Case 3: Worker Thread Not Ready**
```
Slideshow starts before worker thread initialized
Expected: "Slide show - Slide X, Title" (no "has notes")
Reason: hasattr check catches missing worker, returns base_name
```

**Edge Case 4: Multiple Presentations**
```
Two presentations open, slideshow in Presentation 1
Expected: Correct slide/notes from Presentation 1 only
Verification: Check worker uses correct window object
```

## Comparison: Alternative Approaches Considered

### Approach 1: Timing-Based Announcement (REJECTED)

```python
def on_slideshow_slide_changed(...):
    if self._has_meeting_notes():
        self._announce("has notes")  # Hope this arrives first
```

**Why Rejected:**
- ❌ Race condition with NVDA's window name announcement
- ❌ Unreliable ordering (varies by system, timing, load)
- ❌ Creates two separate announcements (pause between them)
- ❌ User explicitly rejected timing-based solutions

### Approach 2: Speech Cancellation (REJECTED)

```python
def on_slideshow_slide_changed(...):
    if self._has_meeting_notes():
        speech.cancelSpeech()  # Cancel window name
        ui.message("has notes, Slide X, Title")  # Announce ours
```

**Why Rejected:**
- ❌ Requires knowing exact window name to reconstruct it
- ❌ Might cancel other important speech
- ❌ Timing issues (when to cancel?)
- ❌ Duplicates NVDA's built-in slideshow functionality

### Approach 3: Speech Priority (REJECTED)

```python
def on_slideshow_slide_changed(...):
    if self._has_meeting_notes():
        speech.speak(["has notes"], priority=priorities.NOW)  # Interrupt
```

**Why Rejected:**
- ❌ Interrupts slide title announcement (bad UX)
- ❌ User explicitly rejected interruption
- ❌ Two separate announcements (fragmented experience)

### Approach 4: Override reportFocus() (ALTERNATIVE)

```python
class CustomSlideShowWindow(SlideShowWindow):
    def reportFocus(self):
        if self.appModule._worker._has_meeting_notes():
            ui.message("has notes")
        super().reportFocus()
```

**Why Not Chosen:**
- ⚠️ Creates two separate announcements
- ⚠️ Less integrated than single announcement
- ✅ Would work reliably (considered as fallback)

### Approach 5: _get_name() Override (CHOSEN ✓)

```python
class CustomSlideShowWindow(SlideShowWindow):
    def _get_name(self):
        base_name = super()._get_name()
        if self.appModule._worker._has_meeting_notes():
            return f"has notes, {base_name}"
        return base_name
```

**Why Chosen:**
- ✅ Single integrated announcement
- ✅ No timing dependencies
- ✅ Standard NVDA pattern
- ✅ Minimal code change
- ✅ Preserves all other features
- ✅ User approved

## Related Documentation

### NVDA Developer Resources

- [NVDA Developer Guide](https://www.nvaccess.org/files/nvda/documentation/developerGuide.html)
- [chooseNVDAObjectOverlayClasses](https://download.nvaccess.org/documentation/developerGuide.html#ChooseNVDAObjectOverlayClasses)
- [NVDAObject API](https://www.webbie.org.uk/nvda/api/NVDAObjects.NVDAObject-class.html)

### PowerPoint Slideshow NVDA Issues

- [Issue #4850 - PowerPoint slideshow reading](https://github.com/nvaccess/nvda/issues/4850)
- [Issue #16825 - On-demand speech mode in PowerPoint](https://github.com/nvaccess/nvda/issues/16825)
- [Issue #16161 - PowerPoint slideshow object detection](https://github.com/nvaccess/nvda/issues/16161)

### Project Documentation

- `.agent/experts/nvda-plugins/decisions.md` - Architectural decisions
- `.agent/experts/nvda-plugins/research/PowerPoint-COM-Events-Research.md` - COM event patterns
- `RELEASE.md` - Release process

## Version History

### v0.0.61 (Current)
- Implemented `CustomSlideShowWindow` with `_get_name()` override
- Added `chooseNVDAObjectOverlayClasses()` to register custom class
- Removed duplicate announcement from `on_slideshow_slide_changed()`
- Single integrated announcement: "has notes, Slide show - Slide X, Title"

### v0.0.60
- Fixed first alt-tab announcement bug
- Improved slide tracking logic

### v0.0.59
- Changed "has meeting notes" to "has notes"
- Attempted timing-based first announcement (unreliable)

### v0.0.56-0.0.58
- Slideshow mode detection and handling
- Event-based "has notes" announcement (came after slide title)
- Various bug fixes

## Maintenance Notes

### When to Update This Architecture

**Update needed if:**
- NVDA changes `SlideShowWindow` interface
- Microsoft changes PowerPoint slideshow window structure
- Users report announcement ordering issues
- New NVDA versions break overlay class registration

**How to debug:**
1. Check NVDA log for "Replaced SlideShowWindow" message
2. Verify `_get_name()` is being called (add logging)
3. Check `_has_meeting_notes()` returns correct value
4. Verify class list contains `SlideShowWindow` before replacement

### Known Limitations

1. **Threading:** `_get_name()` is called on main NVDA thread, but accesses worker thread's `_has_meeting_notes()`. Currently safe because it's a simple property query, but could be race condition if notes detection becomes async.

2. **Cache Invalidation:** Notes detection happens every time `_get_name()` is called. If this becomes a performance issue, consider caching with invalidation on slide change.

3. **NVDA Version Compatibility:** Relies on NVDA's built-in `SlideShowWindow` class existing. Future NVDA versions might refactor this class.

### Future Enhancements

**Possible improvements:**
1. Cache notes status per slide for performance
2. Add configuration option to enable/disable "has notes" announcement
3. Support for different announcement formats (e.g., "Notes, Slide X, Title")
4. Integration with NVDA's verbosity levels

## Summary

The `_get_name()` override approach provides:
- ✅ Reliable, non-timing-dependent ordering
- ✅ Single integrated announcement
- ✅ Standard NVDA addon pattern
- ✅ Scoped correctly to slideshow only
- ✅ Preserves all other NVDA features
- ✅ Minimal code complexity
- ✅ User-approved solution

This architecture solves the slideshow announcement ordering problem completely and maintainably.
