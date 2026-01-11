# NVDA PowerPoint Slideshow Mode Override Research

## Executive Summary

This document provides comprehensive research on how NVDA handles PowerPoint slideshow mode announcements and presents three viable options for implementing the user's requirements: announcing prefixes ("has notes", "Has N comments") BEFORE slide content, and reading ONLY the slide title (not full slide content).

**Recommendation:** Option 2 (TreeInterceptor Override) provides the best balance of reliability, maintainability, and alignment with NVDA architecture. However, Option 3 (Hybrid COM + Overlay) offers a pragmatic fallback if TreeInterceptor customization proves too complex.

---

## Research Findings

### 1. NVDA's SlideShowWindow Architecture

#### Class Definition and Hierarchy

NVDA's built-in `SlideShowWindow` class (from `source/appModules/powerpnt.py`) extends `PaneClassDC` and manages slideshow presentation mode:

```
NVDAObject (base)
    |
IAccessible (interface)
    |
Window (window wrapper)
    |
PaneClassDC (DC pane handling)
    |
SlideShowWindow (PowerPoint slideshow)
    |
CustomSlideShowWindow (our override - partially working)
```

#### Key Properties of SlideShowWindow

| Property | Purpose |
|----------|---------|
| `notesMode` | Boolean flag - True when viewing speaker notes |
| `treeInterceptorClass` | Set to `ReviewableSlideshowTreeInterceptor` |
| `_lastSlideChangeID` | Tracks slide transitions to prevent duplicates |
| `View` | COM SlideShowView object (access to Slide COM object) |
| `currentSlide` | NVDA wrapper object (NOT PowerPoint COM object) |

#### The `_get_name()` Method

NVDA's `SlideShowWindow._get_name()` constructs the announcement:

```python
def _get_name(self):
    # Returns:
    # "Slide show - {slideName}" (normal mode)
    # "Slide show notes - {slideName}" (notes mode)
    # "Slide Show - complete" (end of slideshow)
```

**Critical Finding:** Our `CustomSlideShowWindow._get_name()` override IS working in v0.0.76, but the challenge is accessing slide data (notes/comments) reliably from within this method.

### 2. Slideshow Announcement Flow

#### Complete Event Sequence

```
T0: User presses Space/Arrow/PageDown
    |
T1: NVDA's SlideShowTreeInterceptor.script_slideChange() intercepts
    |
T2: Gesture sent to PowerPoint
    |
T3: PowerPoint advances slide internally
    |
T4: PowerPoint fires SlideShowNextSlide COM event
    |-- Worker thread: on_slideshow_slide_changed() receives event
    |-- Worker thread: Updates cached slide data
    |
T5: PowerPoint window name changes (detected by NVDA)
    |
T6: NVDA's handleSlideChange() triggered
    |
T7: handleSlideChange() calls reportFocus() on SlideShowWindow
    |
T8: reportFocus() queries _get_name() for announcement
    |
T9: NVDA speaks the window name
    |
T10: reportNewSlide() checks autoSayAllOnPageLoad setting
    |
T11: If enabled, sayAll begins reading slide content
```

#### Key Method: `handleSlideChange()`

This method is the central coordinator for slide announcements:

```python
def handleSlideChange(self):
    # 1. Invalidate cached slide data
    # 2. Compare current slide ID against _lastSlideChangeID
    # 3. If changed, clear cached text
    # 4. Call reportFocus() - announces window name
    # 5. Call reportNewSlide() - controls content reading
```

#### Key Method: `reportNewSlide()`

Controls whether slide content is automatically read:

```python
def reportNewSlide(self):
    if config.conf["virtualBuffers"]["autoSayAllOnPageLoad"]:
        # Starts sayAll - reads full slide content
        speechSequence = self.makeTextInfo(...).getSpeechTextForReading()
        speech.speak(speechSequence)
    else:
        # Only reads current line or reports focus position
        pass
```

**Important Configuration:** `virtualBuffers.autoSayAllOnPageLoad` controls automatic content reading!

### 3. Controlling Slide Content Reading

#### Option A: Disable autoSayAllOnPageLoad Globally

Users can disable "Automatic Say All on page load" in Browse Mode settings.

**Pros:**
- Works immediately
- No code changes needed

**Cons:**
- Affects ALL applications, not just PowerPoint
- User may want sayAll in browsers but not PowerPoint
- Doesn't add our prefix

#### Option B: Override reportNewSlide() in Addon

Override the method that triggers content reading:

```python
class CustomSlideShowWindow(SlideShowWindow):
    def reportNewSlide(self):
        # Do nothing - suppress content reading entirely
        pass
```

**Pros:**
- PowerPoint-specific
- Simple implementation

**Cons:**
- Completely suppresses content reading
- User loses ability to hear slide content at all

#### Option C: Custom TreeInterceptor with Modified sayAll

Create a custom TreeInterceptor that:
1. Suppresses automatic sayAll
2. Only speaks title text

**Pros:**
- Full control over what gets read
- Can implement "title only" reading

**Cons:**
- Complex to implement
- Must replicate NVDA's content extraction logic

### 4. COM Integration in Slideshow Mode

#### The View Property Problem

**Discovery from v0.0.61-v0.0.67 attempts:**

- `self.currentSlide` is an NVDA wrapper object, NOT a PowerPoint COM object
- `self.View` exists on `SlideShowWindow` but accessing it from `CustomSlideShowWindow` is problematic
- The worker thread's `_slideshow_window.View.Slide` pattern works because it uses the actual COM object

#### Why Worker Thread Access Works

```python
# WORKER THREAD (works):
slide = self._slideshow_window.View.Slide  # COM object
notes_page = slide.NotesPage
placeholder = notes_page.Shapes.Placeholders(2)
# Successfully reads notes
```

The worker thread receives the actual `SlideShowWindow` COM object via the `SlideShowNextSlide` event parameter.

#### Why Overlay Class Access Fails

```python
# OVERLAY CLASS (fails):
slide = self.View.Slide  # self.View may be None or inaccessible
# Error: View not available
```

The overlay class instance (`CustomSlideShowWindow`) is an NVDA wrapper object with different property bindings.

### 5. Current Implementation Status (v0.0.76)

#### What Works

1. **CustomSlideShowWindow is instantiated** - Logs confirm class replacement
2. **`_get_name()` IS called** - For window name announcement
3. **Worker thread COM access** - Reliably gets slide data
4. **Cached values pattern** - Worker caches `_last_comment_count` and `_last_has_notes`
5. **CustomSlide for normal mode** - `_get_name()` override working with cached values

#### What Doesn't Work

1. **Slideshow `_check_slide_has_notes()`** - Cannot access COM from overlay class
2. **Prefix in slideshow** - Shows "has notes = False" due to COM access failure
3. **Content suppression** - Full slide content still reads after title

---

## Option Analysis

### Option 1: Worker Thread Communication Pattern

**Approach:** Have the overlay class read cached values from the worker thread (similar to normal mode `CustomSlide`).

#### Implementation

```python
class CustomSlideShowWindow(SlideShowWindow):
    def _get_name(self):
        base_name = super()._get_name()

        # Access worker thread's cached values
        app_module = self.appModule
        if app_module and hasattr(app_module, '_worker'):
            worker = app_module._worker
            if worker and worker._initialized:
                prefix_parts = []
                if getattr(worker, '_last_has_notes', False):
                    prefix_parts.append("has notes")
                comment_count = getattr(worker, '_last_comment_count', 0)
                if comment_count > 0:
                    prefix_parts.append(f"Has {comment_count} comment{'s' if comment_count != 1 else ''}")

                if prefix_parts:
                    return f"{', '.join(prefix_parts)}, {base_name}"

        return base_name

    def reportNewSlide(self):
        # Override to suppress full content reading
        # Only announce window name (already includes prefix via _get_name)
        pass  # Do nothing - prevents sayAll
```

#### Technical Feasibility

| Aspect | Assessment |
|--------|------------|
| Prefix before title | HIGH - Same pattern as CustomSlide |
| Only read title | MEDIUM - Requires reportNewSlide override |
| Timing reliability | MEDIUM - Worker must update cache before _get_name() |
| Integration risk | LOW - Minimal changes to existing code |

#### Resource Implications

- **Time:** 2-4 hours implementation, 2-4 hours testing
- **Risk:** Worker cache timing may not be synchronized with NVDA's _get_name() call

#### Advantages

1. Reuses existing worker thread infrastructure
2. Consistent with normal mode pattern
3. Minimal new code

#### Disadvantages

1. **Timing dependency:** Worker must process COM event BEFORE NVDA calls _get_name()
2. **Cache staleness:** First slide may have wrong data if cache not initialized
3. **Doesn't control content reading directly**

---

### Option 2: Custom TreeInterceptor Override (RECOMMENDED)

**Approach:** Create a custom `ReviewableSlideshowTreeInterceptor` subclass that:
1. Suppresses automatic sayAll
2. Reads only slide title
3. Prepends our prefix

#### Implementation

```python
from nvdaBuiltin.appModules.powerpnt import ReviewableSlideshowTreeInterceptor

class CustomSlideshowTreeInterceptor(ReviewableSlideshowTreeInterceptor):
    """Custom slideshow handling with prefix and title-only reading."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._suppress_sayall = True

    def reportNewSlide(self):
        """Override to control what gets announced on slide change."""
        # Get prefix from worker thread
        prefix = self._get_prefix()

        # Get slide title only (not full content)
        title = self._get_slide_title()

        # Build announcement
        parts = []
        if prefix:
            parts.append(prefix)
        if title:
            parts.append(title)
        else:
            parts.append(f"Slide {self._get_slide_number()}")

        # Announce our custom message
        ui.message(", ".join(parts))

        # DO NOT call super().reportNewSlide() - prevents sayAll

    def _get_prefix(self):
        """Get notes/comments prefix from worker."""
        app_module = self.rootNVDAObject.appModule
        if not app_module or not hasattr(app_module, '_worker'):
            return ""

        worker = app_module._worker
        if not worker:
            return ""

        parts = []
        if getattr(worker, '_last_has_notes', False):
            parts.append("has notes")
        comment_count = getattr(worker, '_last_comment_count', 0)
        if comment_count > 0:
            parts.append(f"Has {comment_count} comment{'s' if comment_count != 1 else ''}")

        return ", ".join(parts)

    def _get_slide_title(self):
        """Extract only slide title, not full content."""
        # Use worker's COM access to get title
        app_module = self.rootNVDAObject.appModule
        if app_module and hasattr(app_module, '_worker'):
            worker = app_module._worker
            if worker and worker._slideshow_window:
                try:
                    slide = worker._slideshow_window.View.Slide
                    if slide.Shapes.HasTitle:
                        return slide.Shapes.Title.TextFrame.TextRange.Text.strip()
                except Exception:
                    pass
        return ""

    def _get_slide_number(self):
        """Get current slide number."""
        app_module = self.rootNVDAObject.appModule
        if app_module and hasattr(app_module, '_worker'):
            worker = app_module._worker
            if worker:
                return worker._last_announced_slide
        return 0

# In CustomSlideShowWindow:
class CustomSlideShowWindow(SlideShowWindow):
    treeInterceptorClass = CustomSlideshowTreeInterceptor
```

#### Technical Feasibility

| Aspect | Assessment |
|--------|------------|
| Prefix before title | HIGH - Full control over announcement |
| Only read title | HIGH - We build the message ourselves |
| Timing reliability | HIGH - TreeInterceptor controls entire flow |
| Integration risk | MEDIUM - More complex override |

#### Resource Implications

- **Time:** 4-8 hours implementation, 4-8 hours testing
- **Expertise:** Requires understanding of NVDA TreeInterceptor architecture

#### Advantages

1. **Full control** over what gets announced
2. **No timing issues** - we control the entire announcement flow
3. **Clean separation** - TreeInterceptor handles slideshow, overlay handles window name
4. **NVDA architecture aligned** - This is how NVDA expects customization

#### Disadvantages

1. **Complexity** - TreeInterceptor is a more complex component
2. **NVDA version sensitivity** - May break with NVDA updates
3. **Testing burden** - More scenarios to test

---

### Option 3: Hybrid COM Event + Overlay Pattern

**Approach:** Use COM events (already working) for timing, with overlay class for announcement modification.

#### Implementation

```python
# In PowerPointEventSink.SlideShowNextSlide:
def SlideShowNextSlide(self, slideShowWindow):
    try:
        # Get slide data
        slide_index = slideShowWindow.View.Slide.SlideIndex

        # Update cached values BEFORE NVDA processes
        self._worker._update_slideshow_cache(slideShowWindow)

        # Signal that data is ready
        self._worker._slideshow_data_ready = True

    except Exception as e:
        log.error(f"SlideShowNextSlide error: {e}")

# In PowerPointWorker:
def _update_slideshow_cache(self, slideshow_window):
    """Update cache with slideshow slide data."""
    try:
        slide = slideshow_window.View.Slide

        # Get slide title
        self._slideshow_title = ""
        if slide.Shapes.HasTitle:
            self._slideshow_title = slide.Shapes.Title.TextFrame.TextRange.Text.strip()

        # Get notes status
        self._slideshow_has_notes = self._check_slideshow_notes(slide)

        # Get comment count (if accessible)
        self._slideshow_comment_count = self._get_slideshow_comments(slide)

    except Exception as e:
        log.error(f"Cache update error: {e}")

# In CustomSlideShowWindow:
class CustomSlideShowWindow(SlideShowWindow):
    def _get_name(self):
        """Return prefix + slide title (not full window name)."""
        app_module = self.appModule
        if not app_module or not hasattr(app_module, '_worker'):
            return super()._get_name()

        worker = app_module._worker
        if not worker or not getattr(worker, '_slideshow_data_ready', False):
            return super()._get_name()

        # Build custom announcement
        parts = []

        # Prefix
        if getattr(worker, '_slideshow_has_notes', False):
            parts.append("has notes")
        comment_count = getattr(worker, '_slideshow_comment_count', 0)
        if comment_count > 0:
            parts.append(f"Has {comment_count} comment{'s' if comment_count != 1 else ''}")

        # Slide title
        title = getattr(worker, '_slideshow_title', '')
        slide_num = getattr(worker, '_last_announced_slide', 0)

        if title:
            parts.append(f"Slide {slide_num}, {title}")
        else:
            parts.append(f"Slide {slide_num}")

        return ", ".join(parts)

    def reportNewSlide(self):
        """Suppress automatic content reading."""
        pass  # Do nothing
```

#### Technical Feasibility

| Aspect | Assessment |
|--------|------------|
| Prefix before title | HIGH - _get_name controls output |
| Only read title | HIGH - reportNewSlide suppressed |
| Timing reliability | HIGH - COM event happens before NVDA processes |
| Integration risk | LOW - Builds on existing infrastructure |

#### Resource Implications

- **Time:** 3-6 hours implementation, 3-6 hours testing
- **Risk:** Lower - leverages proven patterns

#### Advantages

1. **Builds on working infrastructure** - Worker thread COM access already works
2. **Timing reliable** - COM event fires before NVDA's window change detection
3. **Pragmatic** - Works within existing architecture
4. **Lower complexity** than TreeInterceptor approach

#### Disadvantages

1. **Dual synchronization** - Worker cache must be ready before _get_name()
2. **Not pure NVDA pattern** - Mixes COM events with overlay classes
3. **Potential race condition** - Though unlikely due to event ordering

---

## Comparison Matrix

| Criterion | Option 1 (Worker Cache) | Option 2 (TreeInterceptor) | Option 3 (Hybrid) |
|-----------|------------------------|---------------------------|-------------------|
| **Prefix reliability** | MEDIUM | HIGH | HIGH |
| **Title-only reading** | MEDIUM | HIGH | HIGH |
| **Timing guarantee** | LOW | HIGH | MEDIUM-HIGH |
| **Implementation effort** | LOW | HIGH | MEDIUM |
| **NVDA alignment** | MEDIUM | HIGH | MEDIUM |
| **Maintenance burden** | LOW | MEDIUM | LOW |
| **NVDA version resilience** | HIGH | MEDIUM | HIGH |
| **Testing complexity** | LOW | HIGH | MEDIUM |

---

## Recommendation

### Primary Recommendation: Option 2 (TreeInterceptor Override)

**Rationale:**
1. **Full control** - Completely controls the announcement flow
2. **NVDA-aligned** - Uses the intended extension point for slideshow customization
3. **No timing dependencies** - The TreeInterceptor IS the announcement mechanism
4. **Clean separation** - Clear responsibility boundaries

**Implementation Priority:**
1. Create `CustomSlideshowTreeInterceptor` class
2. Override `reportNewSlide()` to build custom announcement
3. Use worker thread for COM data access (already working)
4. Register via `CustomSlideShowWindow.treeInterceptorClass`

### Fallback Recommendation: Option 3 (Hybrid)

If TreeInterceptor override proves too complex or breaks with NVDA updates:

1. Extend existing worker thread COM event handling
2. Update cache IMMEDIATELY in `SlideShowNextSlide` handler
3. Override `_get_name()` to use cached values
4. Override `reportNewSlide()` to suppress content reading

---

## Implementation Considerations

### Threading and Timing

**Critical insight from research:** The COM event `SlideShowNextSlide` fires BEFORE NVDA detects the window change. This means:

```
T1: SlideShowNextSlide fires (our code can update cache here)
    |
T2: Worker thread processes event, updates cache
    |
T3: NVDA detects window change
    |
T4: NVDA calls handleSlideChange()
    |
T5: _get_name() called (cache should be ready)
```

**Key:** Update cache synchronously in the COM event handler (runs on worker thread's STA), NOT asynchronously.

### Notes Detection

**Working pattern for slideshow notes:**
```python
def _check_slideshow_notes(self, slide):
    """Check if slide has meeting notes (****markers)."""
    try:
        notes_page = slide.NotesPage
        placeholder = notes_page.Shapes.Placeholders(2)
        if placeholder.HasTextFrame:
            text_frame = placeholder.TextFrame
            if text_frame.HasText:
                notes_text = text_frame.TextRange.Text.strip()
                return '****' in notes_text
    except Exception:
        pass
    return False
```

### Comment Count in Slideshow

**Note:** Comment count may be harder to get in slideshow mode. The `slide.Comments` collection is available on the Slide object:

```python
def _get_slideshow_comments(self, slide):
    """Get comment count for slide."""
    try:
        return slide.Comments.Count
    except Exception:
        return 0
```

### Suppressing Content Reading

**Two approaches:**

1. **Override `reportNewSlide()`** - Prevents sayAll from starting
2. **Clear TreeInterceptor text buffer** - More drastic but complete

Recommended: Override `reportNewSlide()` to do nothing or announce only our custom message.

---

## Risk Assessment

### High Risk Items

| Risk | Mitigation |
|------|------------|
| TreeInterceptor API changes in NVDA | Maintain compatibility tests, pin NVDA version |
| COM event timing variance | Use synchronous cache update, test on slow systems |
| Worker thread not initialized | Add initialization check, fallback to base behavior |

### Medium Risk Items

| Risk | Mitigation |
|------|------------|
| First slide miss | Ensure SlideShowBegin also updates cache |
| Notes mode toggle | Test Ctrl+Shift+S behavior |
| End of slideshow | Handle "complete" state gracefully |

### Low Risk Items

| Risk | Mitigation |
|------|------------|
| Normal mode regression | Existing tests cover this |
| Performance | Cache approach is lightweight |

---

## Testing Strategy

### Unit Test Cases

1. **Slideshow start (F5)** - First slide should have prefix if notes exist
2. **Slide advance (Space)** - Subsequent slides get correct prefix
3. **Slide back (Backspace)** - Previous slide prefix correct
4. **Notes mode (Ctrl+Shift+S)** - Toggle doesn't break prefix
5. **End slideshow (Escape)** - Clean exit, no crashes
6. **Multiple presentations** - Correct data for active slideshow

### Integration Tests

1. **Full presentation run** - 10+ slides with varying notes/comments
2. **Alt-tab during slideshow** - Prefix survives focus loss/gain
3. **NVDA restart during slideshow** - Graceful recovery

### Accessibility Tests

1. **Speech output verification** - Record and analyze actual speech
2. **User acceptance** - Blind user confirms announcement order
3. **Timing verification** - No gaps or interruptions in speech

---

## References

### NVDA Source Code

- [powerpnt.py - SlideShowWindow class](https://github.com/nvaccess/nvda/blob/master/source/appModules/powerpnt.py)
- [browseMode.py - TreeInterceptor base](https://github.com/nvaccess/nvda/blob/master/source/browseMode.py)

### NVDA Documentation

- [NVDA Developer Guide](https://download.nvaccess.org/documentation/developerGuide.html)
- [chooseNVDAObjectOverlayClasses](https://download.nvaccess.org/documentation/developerGuide.html#ChooseNVDAObjectOverlayClasses)

### NVDA Issues (Relevant)

- [Issue #4850 - PowerPoint slideshow reading](https://github.com/nvaccess/nvda/issues/4850)
- [Issue #16825 - On-demand speech mode in PowerPoint](https://github.com/nvaccess/nvda/issues/16825)
- [PR #17488 - Say all on page load setting](https://github.com/nvaccess/nvda/pull/17488)

### Project Documentation

- `.agent/experts/nvda-plugins/slideshow-announcement-architecture.md`
- `.agent/slideshow-notes-announcement-attempts.md`
- `.agent/experts/nvda-plugins/decisions.md`

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0 | 2026-01-05 | Initial research document |

---

## Appendix: NVDA Slideshow Class Hierarchy

```
ReviewCursorManager
    |
SlideShowTreeInterceptor (mixin)
    |
ReviewableSlideshowTreeInterceptor (combines both)
    |
CustomSlideshowTreeInterceptor (our custom class - Option 2)
```

```
NVDAObject
    |
IAccessible
    |
Window
    |
PaneClassDC
    |
SlideShowWindow
    |
CustomSlideShowWindow (our overlay - Options 1, 3)
```
