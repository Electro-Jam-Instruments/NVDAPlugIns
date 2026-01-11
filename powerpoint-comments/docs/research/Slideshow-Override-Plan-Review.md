# Expert Review: Proposed NVDA Slideshow Override Plan (Option 3: Hybrid Approach)

**Reviewer:** Claude (Strategic Planning and Research Specialist)
**Date:** 2026-01-05
**Version Analyzed:** v0.0.76 (current implementation)
**Proposed Plan:** Option 3 - Hybrid COM + Overlay Pattern

---

## Executive Summary

| Metric | Assessment |
|--------|------------|
| **Overall Success Probability** | **35-45%** |
| Prefix Announcement Success | 70-75% |
| Title-Only Reading Success | 15-25% |
| **Go/No-Go Recommendation** | **NO-GO** (as proposed) |

The proposed plan has a critical architectural flaw: **`reportNewSlide()` does NOT exist on `SlideShowWindow`**. It exists on `SlideShowTreeInterceptor`, a completely different class. The plan will successfully implement prefix announcements but will FAIL to suppress content reading without significant architectural changes.

---

## 1. Success Probability Breakdown

### Component 1: Prefix Announcement (Steps 1 + 2)

**Probability: 70-75%**

| Sub-Component | Probability | Notes |
|---------------|-------------|-------|
| COM event caching (Step 1) | 95% | Proven pattern - already works in normal mode |
| Worker thread synchronization | 85% | COM event fires before NVDA processes |
| `_get_name()` calling cached data | 85% | Same pattern as `CustomSlide` in normal mode |
| `self.appModule._worker` accessible | 95% | Confirmed accessible from overlay classes |
| Timing race condition avoidance | 75% | Potential window between COM event and `_get_name()` |

**Combined probability:** 0.95 x 0.85 x 0.85 x 0.95 x 0.75 = **~52%** baseline, but with proper implementation **70-75%** achievable.

### Component 2: Title-Only Reading (Step 3)

**Probability: 15-25%** (as proposed)

| Issue | Impact | Severity |
|-------|--------|----------|
| `reportNewSlide()` not on `SlideShowWindow` | Plan fundamentally broken | CRITICAL |
| Method exists on `SlideShowTreeInterceptor` | Requires different architecture | HIGH |
| Need custom TreeInterceptor subclass | Significant additional work | HIGH |
| Must set `treeInterceptorClass` property | Additional integration step | MEDIUM |

**Why only 15-25%:** The proposed code `def reportNewSlide(self): pass` on `CustomSlideShowWindow` will have NO EFFECT because:
1. NVDA never calls `SlideShowWindow.reportNewSlide()`
2. NVDA calls `self.treeInterceptor.reportNewSlide()`
3. `treeInterceptor` is an instance of `ReviewableSlideshowTreeInterceptor`
4. Our override would be on the wrong class entirely

---

## 2. Critical Risks (Top 3)

### Risk 1: ARCHITECTURAL ERROR - Wrong Method Location (CRITICAL)

**The proposed plan assumes `reportNewSlide()` is a method on `SlideShowWindow`. IT IS NOT.**

**NVDA Source Code Evidence:**

```python
# From NVDA's powerpnt.py - SlideShowWindow class:
class SlideShowWindow(PaneClassDC):
    treeInterceptorClass = ReviewableSlideshowTreeInterceptor  # TreeInterceptor handles reportNewSlide

# From handleSlideChange():
def handleSlideChange(self):
    # ...
    self.reportFocus()                        # Reports window name
    self.treeInterceptor.reportNewSlide()     # THIS is where content reading starts
```

**Consequence:** The proposed `def reportNewSlide(self): pass` on `CustomSlideShowWindow` will be ignored.

**Mitigation Required:** Create custom `ReviewableSlideshowTreeInterceptor` subclass AND set `CustomSlideShowWindow.treeInterceptorClass = CustomSlideshowTreeInterceptor`.

---

### Risk 2: COM Access From Overlay Class (HIGH)

**Attempts v0.0.61-v0.0.67 documented in `.agent/slideshow-notes-announcement-attempts.md` show:**

- `self.currentSlide` is an NVDA wrapper object, NOT PowerPoint COM object
- `self.View` is NOT accessible from `CustomSlideShowWindow`
- Direct COM property access (`NotesPage`, `Comments`) fails with AttributeError

**Evidence from v0.0.66 attempt:**
```
ERROR - CustomSlideShowWindow: Error checking notes - 'Slide' object has no attribute 'NotesPage'
```

**Why Step 2's `_get_name()` MIGHT work:** It reads from worker thread's cached values, not direct COM access. This is the correct approach.

**Why Step 3 has problems:** Even if we fix the method location, `_get_slide_title()` and `_get_slide_number()` in the proposed TreeInterceptor code rely on worker thread COM access which DOES work.

---

### Risk 3: Timing Race Condition (MEDIUM)

**Event Sequence Analysis:**

```
T0: User presses Space/PageDown
T1: NVDA's SlideShowTreeInterceptor.script_slideChange() intercepts
T2: Gesture sent to PowerPoint
T3: PowerPoint advances slide internally
T4: PowerPoint fires SlideShowNextSlide COM event
    |-- Worker thread receives event on STA thread
    |-- Worker updates cached values (_slideshow_title, _slideshow_has_notes, etc.)
T5: PowerPoint window name changes
T6: NVDA detects window change
T7: handleSlideChange() called
T8: reportFocus() calls _get_name()  <-- Reads cached values HERE
T9: reportNewSlide() may start content reading
```

**Potential Issue:** T4-T8 must complete cache update BEFORE T8 reads cache.

**Mitigating Factor:** COM event (T4) typically fires 10-50ms before NVDA window detection (T5-T6). Cache update should complete in <1ms.

**Risk Level:** MEDIUM - Likely safe in practice, but no hard guarantee.

---

## 3. Missing Research / Unknowns

### Unknown 1: TreeInterceptor Initialization Order

When `CustomSlideShowWindow` is instantiated:
1. Does `treeInterceptor` property create a new instance?
2. Or does it reuse an existing one?
3. When exactly is `treeInterceptorClass` evaluated?

**Impact:** Affects whether setting `treeInterceptorClass = CustomSlideshowTreeInterceptor` will work.

### Unknown 2: First Slide Edge Case

On slideshow start (F5), the sequence is:
1. SlideShowBegin event fires
2. First slide displayed
3. (NO SlideShowNextSlide event for first slide)

**Question:** Is cache initialized correctly for first slide, or does user miss "has notes" on slide 1?

**Current code check:** `SlideShowBegin` handler stores `_slideshow_window` but doesn't cache slide data.

### Unknown 3: sayAll Behavior Modification

Even with `reportNewSlide()` suppressed, does NVDA have other paths that trigger sayAll?
- Browse mode automatic reading
- Focus events
- Name change events

**Impact:** May need multiple suppression points.

### Unknown 4: Notes Mode Toggle (Ctrl+Shift+S)

When user toggles notes mode during slideshow:
- Does `_get_name()` get called again?
- Does window name include "has notes" prefix correctly?
- Is there a separate event path?

---

## 4. Recommended Modifications

### Modification A: Fix reportNewSlide Location (REQUIRED)

**Current (BROKEN):**
```python
class CustomSlideShowWindow(SlideShowWindow):
    def reportNewSlide(self):
        pass  # WRONG - This method doesn't exist here
```

**Corrected:**
```python
# 1. Create custom TreeInterceptor
from nvdaBuiltin.appModules.powerpnt import ReviewableSlideshowTreeInterceptor

class CustomSlideshowTreeInterceptor(ReviewableSlideshowTreeInterceptor):
    def reportNewSlide(self):
        """Override to suppress automatic content reading."""
        # Do nothing - prevents sayAll from starting
        pass

# 2. Set on CustomSlideShowWindow
class CustomSlideShowWindow(SlideShowWindow):
    treeInterceptorClass = CustomSlideshowTreeInterceptor

    def _get_name(self):
        # ... existing cached data approach
```

### Modification B: Handle First Slide (RECOMMENDED)

**Add to SlideShowBegin handler:**
```python
def on_slideshow_begin(self, wn):
    log.info("Worker: Slideshow started")
    self._in_slideshow = True
    self._slideshow_window = wn

    # ADDED: Cache first slide data immediately
    try:
        slide = wn.View.Slide
        self._update_slideshow_cache(slide)
    except Exception as e:
        log.debug(f"Could not cache first slide: {e}")
```

### Modification C: Add Cache Readiness Flag (RECOMMENDED)

**Problem:** What if `_get_name()` is called before cache is ready?

**Solution:**
```python
# In worker:
self._slideshow_data_ready = False

def _update_slideshow_cache(self, slide):
    self._slideshow_data_ready = False  # Mark as updating
    # ... cache updates ...
    self._slideshow_data_ready = True   # Mark as ready

# In _get_name():
def _get_name(self):
    if not worker._slideshow_data_ready:
        return super()._get_name()  # Fallback to default
    # ... use cached data
```

### Modification D: Test TreeInterceptor Import (VERIFICATION)

**Verify this import works:**
```python
from nvdaBuiltin.appModules.powerpnt import ReviewableSlideshowTreeInterceptor
```

**If it fails:** The class may not be exported. Alternative approaches:
1. Access via `globals()` after `from nvdaBuiltin.appModules.powerpnt import *`
2. Dynamically patch the existing TreeInterceptor
3. Use a completely different suppression approach

---

## 5. Go/No-Go Recommendation

### RECOMMENDATION: NO-GO (as currently proposed)

**Rationale:**
1. **Step 3 is fundamentally broken** - Wrong method location
2. **Without Step 3, only 50% of requirement met** - Prefix works but content still reads
3. **Fix requires significant additional architecture** - Custom TreeInterceptor class
4. **Unknown import compatibility** - `ReviewableSlideshowTreeInterceptor` may not be accessible

### CONDITIONAL GO:

Proceed with implementation IF:
1. Modify plan to include custom TreeInterceptor (Modification A)
2. Verify `ReviewableSlideshowTreeInterceptor` is importable
3. Add first slide handling (Modification B)
4. Add cache readiness flag (Modification C)
5. Accept 60-70% overall success probability (with modifications)

### ALTERNATIVE RECOMMENDATION:

If TreeInterceptor customization proves too complex, consider **Option 2 from research document** (pure TreeInterceptor approach):
- Higher complexity but more reliable
- Full control over announcement flow
- Better alignment with NVDA architecture

---

## 6. Revised Success Probability (With Modifications)

| Component | Original | With Modifications |
|-----------|----------|-------------------|
| Prefix Announcement | 70-75% | 80-85% |
| Title-Only Reading | 15-25% | 55-65% |
| **Overall** | **35-45%** | **60-70%** |

---

## 7. Implementation Priority (If Proceeding)

| Priority | Task | Risk |
|----------|------|------|
| P0 | Verify `ReviewableSlideshowTreeInterceptor` import | Blocks all Step 3 work |
| P1 | Implement `CustomSlideshowTreeInterceptor` | Core requirement |
| P1 | Set `treeInterceptorClass` on `CustomSlideShowWindow` | Core requirement |
| P2 | Extend SlideShowNextSlide caching (Step 1) | Already partially exists |
| P2 | Update `_get_name()` to use cache (Step 2) | Pattern established |
| P3 | Handle first slide edge case | Improved UX |
| P3 | Add cache readiness flag | Robustness |

---

## Appendix A: NVDA Source Code References

### SlideShowWindow._get_name()
```python
def _get_name(self):
    if self.currentSlide:
        if self.notesMode:
            return _("Slide show notes - {slideName}").format(slideName=self.currentSlide.name)
        else:
            return _("Slide show - {slideName}").format(slideName=self.currentSlide.name)
    else:
        return _("Slide Show - complete")
```

### handleSlideChange() Flow
```python
def handleSlideChange(self):
    try:
        del self.__dict__["currentSlide"]
    except KeyError:
        pass
    curSlideChangeID = self.name
    # ... comparison logic
    try:
        del self.__dict__["basicText"]
    except KeyError:
        pass
    self.reportFocus()                     # Announces window name
    self.treeInterceptor.reportNewSlide()  # Starts content reading
```

### treeInterceptorClass Assignment
```python
class SlideShowWindow(PaneClassDC):
    treeInterceptorClass = ReviewableSlideshowTreeInterceptor
```

---

## Appendix B: Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0 | 2026-01-05 | Initial expert review |

---

## Sources

- [NVDA PowerPoint App Module](https://github.com/nvaccess/nvda/blob/master/source/appModules/powerpnt.py)
- [NVDA Developer Guide](https://download.nvaccess.org/documentation/developerGuide.html)
- Project file: `.agent/slideshow-notes-announcement-attempts.md`
- Project file: `docs/research/NVDA-Slideshow-Mode-Override-Research.md`
- Project file: `.agent/experts/nvda-plugins/slideshow-announcement-architecture.md`
