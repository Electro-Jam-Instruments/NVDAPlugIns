# Slideshow "Has Notes" Announcement - All Attempts (v0.0.61-v0.0.67)

## Goal
Announce "has notes" BEFORE slide number/title during PowerPoint slideshow presentations.

Example desired output:
- WITH notes: "has notes, Slide show - Slide 1, Title"
- WITHOUT notes: "Slide show - Slide 1, Title"

## Requirements
- NO timing dependencies (must be architecture-based)
- Single announcement (not separate speech events)
- Must NOT interrupt NVDA's normal reading
- Must work reliably across slide changes

---

## Attempt 1: `_get_name()` Override (v0.0.61)

### Approach
Override `_get_name()` method in `CustomSlideShowWindow` to prepend "has notes, " to window name.

### Implementation
```python
class CustomSlideShowWindow(SlideShowWindow):
    def _get_name(self):
        base_name = super()._get_name()
        if hasattr(self.appModule, '_worker') and self.appModule._worker:
            if self.appModule._worker._has_meeting_notes():
                return f"has notes, {base_name}"
        return base_name
```

### Result
**FAILED** - `_get_name()` was NEVER called by NVDA during slideshow.

### Logs Evidence
```
chooseNVDAObjectOverlayClasses: AFTER replacement - clsList=['CustomSlideShowWindow', ...]
# NO logs from _get_name() at all
```

### Analysis
NVDA slideshow uses `handleSlideChange()` → `reportFocus()` path, NOT the normal focus event flow that calls `_get_name()`.

---

## Attempt 2: Add `reportFocus()` Override (v0.0.62)

### Approach
Override `reportFocus()` as the actual entry point for slideshow announcements, keep `_get_name()` as fallback.

### Implementation
```python
def reportFocus(self):
    has_notes = False
    if hasattr(self.appModule, '_worker') and self.appModule._worker:
        has_notes = self.appModule._worker._has_meeting_notes()

    if has_notes:
        base_name = super()._get_name()
        import ui
        ui.message(f"has notes, {base_name}")
    else:
        super().reportFocus()
```

### Result
**PARTIAL** - `reportFocus()` WAS called, but `has_notes` always returned `False`.

### Logs Evidence
```
CustomSlideShowWindow.reportFocus() CALLED
CustomSlideShowWindow.reportFocus(): has_notes = False  # WRONG - slide HAS notes
```

### Analysis
Threading issue - main thread calling worker thread's `_has_meeting_notes()` created race condition. Worker thread data not synchronized with main thread's slideshow state.

---

## Attempt 3: Direct Slide Access via `self.currentSlide` (v0.0.63)

### Approach
Access slide notes directly from `self.currentSlide` instead of worker thread to eliminate threading issues.

### Implementation
```python
def _check_slide_has_notes(self):
    if not self.currentSlide:
        return False

    # Access notes directly
    notes_page = self.currentSlide.NotesPage
    placeholder = notes_page.Shapes.Placeholders(2)
    # ... check for **** markers
```

### Result
**FAILED** - `AttributeError: 'Slide' object has no attribute 'NotesPage'`

### Logs Evidence
```
ERROR - CustomSlideShowWindow: Error checking notes - 'Slide' object has no attribute 'NotesPage'
```

### Analysis
`self.currentSlide` during slideshow is NOT a regular PowerPoint Slide COM object. It's a different type without direct NotesPage access.

---

## Attempt 4: Access via `currentSlide.Parent` (v0.0.64)

### Approach
Assumed `currentSlide` is a `SlideShowView.Slide` object. Access underlying Slide via `.Parent` property.

### Implementation
```python
def _check_slide_has_notes(self):
    if not self.currentSlide:
        return False

    # During slideshow, currentSlide.Parent gives us the actual Slide object
    slide = self.currentSlide.Parent
    notes_page = slide.NotesPage
    # ... check for **** markers
```

### Result
**FAILED** - `AttributeError: 'Slide' object has no attribute 'Parent'`

### Logs Evidence
```
ERROR - CustomSlideShowWindow: Error checking notes - 'Slide' object has no attribute 'Parent'
```

### Analysis
Assumption was wrong - the Slide object itself doesn't have a Parent attribute.

---

## Attempt 5: Add Type Diagnostics (v0.0.65)

### Approach
Add comprehensive diagnostics to discover what `self.currentSlide` actually is.

### Implementation
```python
current_type = type(self.currentSlide).__name__
log.info(f"CustomSlideShowWindow: currentSlide type = {current_type}")

attrs = [attr for attr in dir(self.currentSlide) if not attr.startswith('_')]
log.info(f"CustomSlideShowWindow: Available attributes = {attrs[:20]}")
```

### Result
**DISCOVERY** - `self.currentSlide` is an NVDA wrapper object, NOT a PowerPoint COM object!

### Logs Evidence
```
CustomSlideShowWindow: currentSlide type = Slide  # NVDA object type
CustomSlideShowWindow: Available attributes = ['APIClass', 'appModule', 'TextInfo', 'actionCount', ...]
# These are NVDA properties, NOT PowerPoint COM properties
```

### Analysis
**Critical Finding:** `self.currentSlide` has NVDA properties (`appModule`, `TextInfo`, `APIClass`) not COM properties (`Parent`, `NotesPage`, `SlideIndex`). Cannot use it to access PowerPoint COM object directly.

---

## Attempt 6: Access via `self.View.Slide` (v0.0.66)

### Approach
Based on worker thread pattern at line 772 (`self._slideshow_window.View.Slide`), access PowerPoint Slide COM object via `self.View.Slide`.

### Implementation
```python
def _check_slide_has_notes(self):
    if not hasattr(self, 'View') or not self.View:
        log.debug("CustomSlideShowWindow: No View available")
        return False

    slide = self.View.Slide  # self.View is SlideShowView COM object
    notes_page = slide.NotesPage
    # ... check for **** markers
```

### Result
**FAILED** - Early return, NO logs from inside method except `has_notes = False`.

### Logs Evidence
```
CustomSlideShowWindow.reportFocus() CALLED
CustomSlideShowWindow.reportFocus(): has_notes = False
# NO logs from _check_slide_has_notes() - early return at View check
```

### Analysis
Either `hasattr(self, 'View')` returns `False` OR `self.View` is `None`. The View property is not accessible on `CustomSlideShowWindow`.

---

## Attempt 7: Debug View Property (v0.0.67) - CURRENT

### Approach
Add diagnostics to determine:
1. Does `View` property exist?
2. If yes, what type is it and is it None?
3. If no, what view-related properties ARE available?

### Implementation
```python
has_view = hasattr(self, 'View')
log.info(f"CustomSlideShowWindow: hasattr(self, 'View') = {has_view}")

if has_view:
    view_value = self.View
    view_type = type(view_value).__name__
    log.info(f"CustomSlideShowWindow: self.View type = {view_type}, is_none = {view_value is None}")
else:
    attrs = [attr for attr in dir(self) if 'view' in attr.lower()]
    log.info(f"CustomSlideShowWindow: Attributes containing 'view' = {attrs}")
```

### Result
**PENDING USER TESTING**

### Expected Outcome
Logs will reveal why `View` isn't accessible and what the correct property path is.

---

## Key Learnings

### What We Know Works
1. **Worker thread pattern:** `self._slideshow_window.View.Slide` successfully accesses slide notes (line 772)
2. **CustomSlideShowWindow IS instantiated:** Logs confirm class replacement and `__init__()` calls
3. **reportFocus() IS called:** Entry point for announcements is correct

### What Doesn't Work
1. **`self.currentSlide`** - NVDA wrapper, not PowerPoint COM object
2. **`self.currentSlide.Parent`** - Property doesn't exist
3. **`self.View`** (v0.0.66) - Property not accessible (reason unknown)

### Remaining Questions (v0.0.67 will answer)
1. Why is `View` property not accessible on `CustomSlideShowWindow` when it works on worker's `_slideshow_window`?
2. Is there a different property name for accessing SlideShowView from the overlay class?
3. Do we need to access the COM object through a different NVDA property path?

---

## Architecture Context

### NVDA Slideshow Flow
```
User changes slide
  ↓
PowerPoint fires SlideShowNextSlide event
  ↓
Worker thread: on_slideshow_slide_changed()
  ↓
NVDA: handleSlideChange()
  ↓
CustomSlideShowWindow.reportFocus()  ← Our entry point
  ↓
Need: Access slide.NotesPage here
```

### Object Hierarchy (Expected)
```
CustomSlideShowWindow (NVDA overlay class)
  ├── View (SlideShowView COM object) ← Currently not accessible
  │     └── Slide (PowerPoint Slide COM object)
  │           └── NotesPage ← What we need
  └── currentSlide (NVDA wrapper object) ← NOT usable for COM access
```

---

## Next Steps (After v0.0.67 diagnostics)

### If `View` exists but is None:
- Timing issue - View not initialized when `reportFocus()` called?
- Need to cache View reference during `__init__()`?

### If `View` doesn't exist:
- Find correct property name from diagnostic attrs list
- May need to access via different NVDA property chain
- Investigate if we need `IAccessible` or other NVDA interface

### If completely blocked:
- Consider alternative: Pass slide index to worker thread for notes check
- Worker thread has working COM access pattern
- Main thread just requests result from worker

---

## File Reference
All code changes in: `powerpoint-comments/addon/appModules/powerpnt.py`
- CustomSlideShowWindow class: Lines 1050-1199
- Worker thread pattern (working): Line 772
