# Announcement Patterns Reference

Patterns for modifying NVDA announcements in PowerPoint, particularly for controlling announcement order.

## PowerPoint View Modes

This document covers announcements in both PowerPoint view modes:

- **Normal view** - Where you edit slides (slide thumbnails on left, main slide in center, notes below)
- **Slideshow view** - Full-screen presentation mode (F5 to start)

Each mode has different announcement patterns and override points.

## Problem: Announcement Order (Normal View)

When navigating slides in Normal view, NVDA announces in this order:
1. NVDA built-in: "Slide 1, Title"
2. Our addon: "Has 1 comment"
3. Our addon: "has notes"

**Desired order:**
1. "has notes"
2. "Has 1 comment"
3. "Slide 1, Title"

## Solutions

### Option 1: Intercept Focus Event (Recommended)

Override `event_gainFocus` to announce BEFORE calling `nextHandler()`.

```python
def event_gainFocus(self, obj, nextHandler):
    if self._is_slide_object(obj):
        # Announce our info FIRST
        parts = []
        if has_notes:
            parts.append("has notes")
        if comments_count > 0:
            parts.append(f"Has {comments_count} comment{'s' if comments_count != 1 else ''}")

        if parts:
            ui.message(", ".join(parts))

    # THEN let NVDA announce slide normally
    nextHandler()
```

**Advantages:**
- Clean interception point
- Guaranteed order (our announcement → NVDA's announcement)
- No timing dependencies

### Option 2: Overlay Class with _get_name()

Override `_get_name()` in overlay class to prepend info to name.

```python
class CustomSlide(Slide):
    def _get_name(self):
        base_name = super()._get_name()
        prefix_parts = []

        if has_notes:
            prefix_parts.append("has notes")
        if comments_count > 0:
            prefix_parts.append(f"Has {comments_count} comments")

        if prefix_parts:
            return f"{', '.join(prefix_parts)}, {base_name}"
        return base_name
```

**Advantages:**
- Integrated with NVDA's name resolution
- Lazy evaluation - queries data when needed

**Current implementation:** Uses this pattern for slideshow mode.

### Option 3: Cancel and Replace

Cancel NVDA's announcement and replace with custom.

```python
def event_gainFocus(self, obj, nextHandler):
    if self._is_slide_object(obj):
        cancelSpeech()
        # Build complete announcement ourselves
        ui.message(complete_announcement)
        return  # Don't call nextHandler
```

**Disadvantages:**
- Must reproduce ALL of NVDA's announcement logic
- Risk of missing verbosity settings
- Higher maintenance burden

## Implementation Status

| Mode | Pattern Used | Status |
|------|--------------|--------|
| Slideshow | Overlay `_get_name()` | Working (v0.0.76+) |
| Normal mode | Worker thread announcement | Working but wrong order |
| Normal mode fix | `event_gainFocus` interception | Planned |

## Identifying Slide Objects

To intercept slide focus events, need to identify slide objects:

```python
def _is_slide_object(self, obj):
    # Options to investigate:
    # - obj.role == controlTypes.Role.PANE in slide area
    # - obj.windowClassName specific to slides
    # - Check parent chain for slide container
    pass
```

**Note:** Requires logging to discover actual slide object properties during navigation.

## Related Patterns

### Comment Card Reformatting

Already using cancel-and-replace for comment cards:

```python
def event_gainFocus(self, obj, nextHandler):
    if obj.UIAElement.CurrentClassName == "CommentCard":
        cancelSpeech()
        # Reformat and announce
        ui.message(reformatted_text)
        return  # Don't call nextHandler
```

### Slideshow TreeInterceptor

For slideshow announcements, override `reportNewSlide()` on `ReviewableSlideshowTreeInterceptor`:

```python
class CustomTreeInterceptor(ReviewableSlideshowTreeInterceptor):
    def reportNewSlide(self, suppressSayAll=False):
        # Custom slideshow announcement
        info = self.selection
        info.expand(textInfos.UNIT_LINE)
        speech.speakTextInfo(info, ...)
```

See [slideshow-override.md](./slideshow-override.md) for details.
