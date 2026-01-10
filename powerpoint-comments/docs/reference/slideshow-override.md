# Slideshow Override Reference

How to customize what NVDA announces during PowerPoint slideshow mode.

**Version:** Documented from v0.0.78 implementation (January 2026)

## The Problem We Solved

**Goal:** In slideshow mode, announce:
- "has notes" prefix (if slide has meeting notes)
- "Has N comments" prefix (if slide has comments)
- Slide title only

**NOT announce:** Full slide content (every shape, text box, bullet point)

## NVDA's Slideshow Class Hierarchy

```
┌─────────────────────────────────────────────────────────┐
│                    SlideShowWindow                       │
│  (NVDAObject representing the slideshow window)          │
│                                                          │
│  Key methods:                                            │
│  - _get_name(): Returns window name (lazy evaluation)   │
│  - treeInterceptorClass: Which TreeInterceptor to use   │
└─────────────────────────────────────────────────────────┘
                            │ creates
                            ▼
┌─────────────────────────────────────────────────────────┐
│          ReviewableSlideshowTreeInterceptor              │
│  (Browse mode handler - controls content reading)        │
│                                                          │
│  Key methods:                                            │
│  - reportNewSlide(): Controls slide change announcement │
│                      THIS IS THE KEY METHOD              │
└─────────────────────────────────────────────────────────┘
```

## Critical Insight: Two Separate Announcement Points

| Method | Class | Controls |
|--------|-------|----------|
| `_get_name()` | SlideShowWindow | Window name (slide title with our prefix) |
| `reportNewSlide()` | TreeInterceptor | Content reading (sayAll vs first line) |

**CRITICAL:** `reportNewSlide()` lives on the **TreeInterceptor**, NOT on SlideShowWindow!

## Implementation Pattern

### Step 1: Create Custom TreeInterceptor

```python
from nvdaBuiltin.appModules.powerpnt import ReviewableSlideshowTreeInterceptor

class CustomSlideshowTreeInterceptor(ReviewableSlideshowTreeInterceptor):
    """Override content reading behavior."""

    def reportNewSlide(self, suppressSayAll: bool = False):
        """Announce title only, skip full content reading."""
        import textInfos
        import speech
        import controlTypes

        info = self.selection
        if not info.isCollapsed:
            speech.speakPreselectedText(info.text)
        else:
            info.expand(textInfos.UNIT_LINE)
            speech.speakTextInfo(
                info,
                reason=controlTypes.OutputReason.CARET,
                unit=textInfos.UNIT_LINE
            )
```

### Step 2: Create Custom SlideShowWindow

```python
from nvdaBuiltin.appModules.powerpnt import SlideShowWindow

class CustomSlideShowWindow(SlideShowWindow):
    """Override window name to include prefix."""

    # CRITICAL: Point to our custom TreeInterceptor
    treeInterceptorClass = CustomSlideshowTreeInterceptor

    def _get_name(self):
        """Return window name with notes/comments prefix."""
        prefix_parts = []
        if worker._slideshow_has_notes:
            prefix_parts.append("has notes")
        if worker._slideshow_comment_count > 0:
            prefix_parts.append(f"Has {worker._slideshow_comment_count} comments")

        title = worker._slideshow_title or "Slide"
        if prefix_parts:
            return f"{', '.join(prefix_parts)}, {title}"
        return title
```

### Step 3: Register in chooseNVDAObjectOverlayClasses

```python
def chooseNVDAObjectOverlayClasses(self, obj, clsList):
    super().chooseNVDAObjectOverlayClasses(obj, clsList)

    if SlideShowWindow in clsList:
        idx = clsList.index(SlideShowWindow)
        clsList[idx] = CustomSlideShowWindow
```

## Event Flow on Slide Change

```
1. PowerPoint advances slide
   │
2. COM Event fires (SlideShowNextSlide)
   │  └── Worker thread caches: title, notes, comments
   │
3. NVDA detects window change
   │
4. handleSlideChange() calls treeInterceptor.reportNewSlide()
   │
5. reportNewSlide() speaks first line only
   │
6. User hears: "has notes, Has 2 comments, Slide Title"
```

## Why This Works

1. **COM events fire BEFORE NVDA**: Worker thread has time to cache data
2. **Lazy _get_name()**: Fetched when needed, reads fresh cached data
3. **treeInterceptorClass**: Links SlideShowWindow to our TreeInterceptor
4. **reportNewSlide() override**: Replaces sayAll with single line

## First Slide Edge Case

**Problem:** `SlideShowNextSlide` event does NOT fire for the first slide when slideshow starts.

**Solution:** Cache first slide data in `SlideShowBegin` event handler:

```python
def on_slideshow_begin(self, wn):
    """Called when slideshow starts."""
    self._in_slideshow = True
    self._slideshow_window = wn

    # CRITICAL: Cache first slide data immediately
    # SlideShowNextSlide does NOT fire for slide 1
    try:
        self._cache_slideshow_slide_data(wn)
        log.info("Cached first slide data on slideshow begin")
    except Exception as e:
        log.error(f"Error caching first slide: {e}")
```

This ensures the first slide has the correct prefix announcement.

## Accessing Worker from Overlay Class - The Actual Pattern

The `_get_name()` examples reference `worker._slideshow_has_notes`. The actual implementation uses a **module-level global**:

```python
# At module level (top of powerpnt.py)
_current_app_module = None

class AppModule(AppModule):
    def __init__(self, *args, **kwargs):
        global _current_app_module
        super().__init__(*args, **kwargs)
        _current_app_module = self  # Store reference
        self._worker = PowerPointWorker()
        self._worker.start()

# In overlay class
class CustomSlideShowWindow(SlideShowWindow):
    def _get_name(self):
        global _current_app_module
        if _current_app_module and _current_app_module._worker:
            worker = _current_app_module._worker
            # Now access worker._slideshow_has_notes, etc.
```

**Why module-level global?** The overlay class `_get_name()` receives `self` (the NVDAObject), not the AppModule. Using `getAppModuleForNVDAObject(self)` is an alternative but the global is simpler and what the actual code uses.

## Testing Checklist

- [ ] Slide title announces on slide change
- [ ] Prefix (notes/comments) appears before title
- [ ] Full slide content suppressed (no sayAll)
- [ ] **First slide works** (uses SlideShowBegin caching)
- [ ] Going backwards works
- [ ] Ending slideshow (Escape) works without errors
