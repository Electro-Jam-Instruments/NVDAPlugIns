# Feature Update Implementation Plan

**Created:** 2026-01-10
**Version:** 0.0.78 (baseline)
**Status:** Analysis Complete - Ready for Implementation

---

## Executive Summary

This document provides a comprehensive implementation plan for four feature updates to the PowerPoint Comments NVDA addon. Each task includes current state analysis, implementation approach, risk assessment, and testing procedures.

---

## Task 1: Fix Off-By-One Slide Number Bug

### Problem Statement
The slide number being announced appears to be off by one.

### Current State Analysis

**Where slide numbers are determined:**

| Location | File:Line | Purpose |
|----------|-----------|---------|
| `_get_current_slide_index()` | powerpnt.py:825-832 | Gets 1-based index via COM |
| `_cache_slideshow_slide_data()` | powerpnt.py:709-764 | Caches `slide.SlideIndex` |
| `CustomSlide._get_name()` | powerpnt.py:1369-1415 | Returns `super()._get_name()` |
| `CustomSlideShowWindow._get_name()` | powerpnt.py:1287-1345 | Uses `self.currentSlide.name` |
| `_announce_slide_comments()` | powerpnt.py:1037-1064 | Uses `_get_current_slide_index()` |

**Key Code Paths:**

1. **Normal Mode (CustomSlide):**
   ```python
   # Line 1381 - Gets base name from NVDA's built-in Slide class
   base_name = super()._get_name()
   # Format: "Slide N (Title)" where N comes from NVDA's Slide class
   ```

2. **Slideshow Mode (CustomSlideShowWindow):**
   ```python
   # Lines 1330-1336 - Uses cached title or currentSlide.name
   cached_title = getattr(worker, '_slideshow_title', '')
   if cached_title:
       slide_part = cached_title  # Title without "Slide N"
   elif self.currentSlide:
       slide_part = self.currentSlide.name  # NVDA wrapper's name
   ```

3. **COM-based slide index:**
   ```python
   # Line 830 - COM API uses 1-based indexing
   return window.View.Slide.SlideIndex
   ```

**Analysis:**
- PowerPoint's COM API uses **1-based** `SlideIndex` (confirmed in docs)
- NVDA's built-in `Slide._get_name()` should return correct "Slide N" based on `SlideIndex`
- The off-by-one could be in:
  1. NVDA's built-in class (unlikely - widely used)
  2. Our caching of `SlideIndex` before/after navigation
  3. Timing issue where old slide index is used before navigation completes

**Most Likely Cause:**
The worker thread's `_last_announced_slide` tracking or the slide change event timing. Need to verify if the bug is:
- In the slide number NVDA announces (from built-in class)
- In the comment count for wrong slide (our caching)
- In slideshow vs normal mode or both

### Investigation Steps (Before Implementation)

1. Add debug logging to confirm where the number comes from:
   ```python
   # In CustomSlide._get_name()
   log.info(f"CustomSlide base_name='{base_name}', worker slide index={worker._get_current_slide_index()}")
   ```

2. Test in both modes:
   - Normal view: Navigate with arrow keys
   - Slideshow: Navigate with arrow keys

3. Compare NVDA's announced slide number vs actual PowerPoint slide

### Potential Fix Approaches

**If caching issue:**
- Review timing of `_last_announced_slide` updates
- Ensure `on_slide_changed_event()` receives correct index

**If timing issue:**
- Ensure COM event fires AFTER slide change completes
- Verify `WindowSelectionChange` provides correct slide, not previous

### Risk Assessment

| Risk | Likelihood | Impact | Mitigation |
|------|------------|--------|------------|
| Timing race condition | Medium | High | Use lazy evaluation pattern |
| Break existing functionality | Low | High | Test all navigation scenarios |
| COM event order issues | Medium | Medium | Verify with debug logging first |

**Reference Pitfalls:**
- Pitfall 9: Worker Thread Data Stale at Event Time
- Pitfall 12: Using ActiveWindow with Multiple Presentations

---

## Task 2: Update Comment Date/Time Display

### Problem Statement
**Current:** Date/time is completely removed from comments
**Required:**
- If comment is within last 7 days: Show as "N days ago"
- If older than 7 days: Don't show date/time at all

### Current State Analysis

**Where comment formatting happens:**

| Location | File:Line | Pattern |
|----------|-----------|---------|
| Thread cards (cardRoot_) | powerpnt.py:1649-1674 | Extracts author, uses description |
| Reply comments (postRoot_) | powerpnt.py:1676-1715 | Parses "Comment by Author on..." |

**Current Comment Name Format (from UIA):**
- Thread: `"Comment thread started by Author, with N replies. {date} {time}"`
- Reply: `"Comment by Author on Month Day, Year, Time"`

**Current Processing:**
```python
# Thread cards (line 1654-1659):
if " started by " in name_normalized:
    author_part = name_normalized.split(" started by ", 1)[1]
    if ", with " in author_part:
        author = author_part.split(", with ")[0]
# Result: "Author: description" - date is in name but not used

# Reply comments (line 1682-1694):
if name_normalized.startswith("Comment by "):
    after_prefix = name_normalized[11:]  # Skip "Comment by "
    if " on " in after_prefix:
        author = after_prefix.split(" on ", 1)[0]
# Result: "Author: description" - date after " on " is discarded
```

### Implementation Plan

**Step 1: Create date parsing helper function**

Location: Add after line 1529 (in `event_gainFocus` method area)

```python
def _parse_comment_date(self, name_normalized):
    """Parse date from comment name and calculate days ago.

    Args:
        name_normalized: Whitespace-normalized comment name

    Returns:
        tuple: (datetime object or None, "N days ago" string or "")
    """
    import re
    from datetime import datetime, timedelta

    # Pattern for "on Month Day, Year, Time" format
    # Example: "on January 5, 2026, 2:30 PM"
    date_pattern = r' on (\w+ \d{1,2}, \d{4})'
    match = re.search(date_pattern, name_normalized)

    if not match:
        return None, ""

    try:
        date_str = match.group(1)
        comment_date = datetime.strptime(date_str, "%B %d, %Y")
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        delta = today - comment_date.replace(hour=0, minute=0, second=0, microsecond=0)
        days = delta.days

        if days < 0:
            return comment_date, "today"  # Future date edge case
        elif days == 0:
            return comment_date, "today"
        elif days == 1:
            return comment_date, "1 day ago"
        elif days <= 7:
            return comment_date, f"{days} days ago"
        else:
            return comment_date, ""  # Older than 7 days - no date shown

    except ValueError as e:
        log.debug(f"Could not parse comment date: {e}")
        return None, ""
```

**Step 2: Modify thread card formatting (lines 1649-1674)**

```python
if is_comment_card:
    is_resolved = name_normalized.startswith("Resolved ")
    author = ""

    if " started by " in name_normalized:
        author_part = name_normalized.split(" started by ", 1)[1]
        if ", with " in author_part:
            author = author_part.split(", with ")[0]
        else:
            author = author_part

    # NEW: Parse date and get "N days ago" string
    _, days_ago = self._parse_comment_date(name_normalized)

    if author and description:
        # Skip cancelSpeech after slide navigation
        if not getattr(self, '_just_navigated', False):
            speech.cancelSpeech()
        else:
            self._just_navigated = False
            log.info("Skipped cancelSpeech - letting slide title finish")

        # Build formatted output with optional date
        if is_resolved:
            if days_ago:
                formatted = f"Resolved - {author}, {days_ago}: {description}"
            else:
                formatted = f"Resolved - {author}: {description}"
        else:
            if days_ago:
                formatted = f"{author}, {days_ago}: {description}"
            else:
                formatted = f"{author}: {description}"

        ui.message(formatted)
        log.info(f"Comment reformatted: {formatted[:80]}")
        return
```

**Step 3: Modify reply comment formatting (lines 1676-1715)**

```python
elif is_reply_comment:
    author = ""
    is_task_status = False

    if name_normalized.startswith("Task updated by "):
        after_prefix = name_normalized[16:]
        if " on " in after_prefix:
            author = after_prefix.split(" on ", 1)[0]
        is_task_status = True
    elif name_normalized.startswith("Comment by "):
        after_prefix = name_normalized[11:]
        if " on " in after_prefix:
            author = after_prefix.split(" on ", 1)[0]

    # NEW: Parse date and get "N days ago" string
    _, days_ago = self._parse_comment_date(name_normalized)

    if author and description:
        if not getattr(self, '_just_navigated', False):
            speech.cancelSpeech()
        else:
            self._just_navigated = False
            log.info("Skipped cancelSpeech - letting slide title finish")

        if is_task_status:
            status_text = description.replace(" a task", " task")
            if days_ago:
                formatted = f"{author}, {days_ago} - {status_text}"
            else:
                formatted = f"{author} - {status_text}"
        else:
            if days_ago:
                formatted = f"{author}, {days_ago}: {description}"
            else:
                formatted = f"{author}: {description}"

        ui.message(formatted)
        log.info(f"Reply/Status reformatted: {formatted[:80]}")
        return
```

### Risk Assessment

| Risk | Likelihood | Impact | Mitigation |
|------|------------|--------|------------|
| Date format varies by locale | Medium | Medium | Test with different Windows locales |
| Parsing fails silently | Low | Low | Fallback returns empty string |
| Non-breaking space in date | Medium | Medium | Already normalizing whitespace |

**Reference Pitfalls:**
- Pitfall 10: Non-Breaking Spaces in PowerPoint Text (v0.0.42)

### Testing Plan

1. Create comment today - expect "today"
2. Create comment yesterday - expect "1 day ago"
3. Create comment 5 days ago - expect "5 days ago"
4. Create comment 8 days ago - expect no date (just "Author: text")
5. Test with both thread (cardRoot_) and reply (postRoot_) comments
6. Test with resolved comments

---

## Task 3: Change Announcement Order

### Problem Statement
**Current order:** [slide info], [has notes], [has X comments], [title]
**Required order:** [slide info], [has X comments], [has notes info], [title]

### Current State Analysis

**Where order is determined:**

| Location | File:Line | Current Order |
|----------|-----------|---------------|
| `CustomSlide._get_name()` | powerpnt.py:1395-1407 | has_notes, comments, base_name |
| `CustomSlideShowWindow._get_name()` | powerpnt.py:1315-1340 | has_notes, comments, slide_part |
| `_announce_slide_comments()` | powerpnt.py:1043-1056 | has_notes, comments |

**Current Code (CustomSlide._get_name()):**
```python
# Lines 1395-1407
prefix_parts = []

has_notes = getattr(worker, '_last_has_notes', False)
if has_notes:
    prefix_parts.append("has notes")  # Added first

comment_count = getattr(worker, '_last_comment_count', 0)
if comment_count > 0:
    if comment_count == 1:
        prefix_parts.append("Has 1 comment")  # Added second
    else:
        prefix_parts.append(f"Has {comment_count} comments")
```

**Understanding Announcement Structure:**

The current format `super()._get_name()` returns: `"Slide N (Title)"` or `"Slide N"` if no title.

Full announcement example: `"has notes, Has 2 comments, Slide 3 (Project Timeline)"`

The user's description breaks this down as:
- [slide info] = "Slide 3"
- [has notes] = "has notes"
- [has X comments] = "Has 2 comments"
- [title] = "(Project Timeline)"

However, the current code treats "Slide N (Title)" as one unit from the parent class.

### Implementation Plan

**Option A: Simple Reorder (Swap notes and comments)**

Change in 3 locations:

**1. CustomSlide._get_name() (lines 1395-1407):**
```python
prefix_parts = []

# REORDERED: Comments first, then notes
comment_count = getattr(worker, '_last_comment_count', 0)
if comment_count > 0:
    if comment_count == 1:
        prefix_parts.append("Has 1 comment")
    else:
        prefix_parts.append(f"Has {comment_count} comments")

has_notes = getattr(worker, '_last_has_notes', False)
if has_notes:
    prefix_parts.append("has notes")
```

**2. CustomSlideShowWindow._get_name() (lines 1315-1327):**
```python
prefix_parts = []

# REORDERED: Comments first, then notes
comment_count = getattr(worker, '_slideshow_comment_count', 0)
if comment_count > 0:
    if comment_count == 1:
        prefix_parts.append("Has 1 comment")
    else:
        prefix_parts.append(f"Has {comment_count} comments")

has_notes = getattr(worker, '_slideshow_has_notes', False)
if has_notes:
    prefix_parts.append("has notes")
```

**3. _announce_slide_comments() (lines 1043-1055):**
```python
# Build prefix for Comments pane navigation
prefix_parts = []

# REORDERED: Comments first, then notes
if self._last_comment_count > 0:
    if self._last_comment_count == 1:
        prefix_parts.append("Has 1 comment")
    else:
        prefix_parts.append(f"Has {self._last_comment_count} comments")

if self._last_has_notes:
    prefix_parts.append("has notes")
```

**Result after Option A:**
`"Has 2 comments, has notes, Slide 3 (Project Timeline)"`

### Risk Assessment

| Risk | Likelihood | Impact | Mitigation |
|------|------------|--------|------------|
| User confusion from change | Low | Low | Document in changelog |
| Inconsistent behavior | Low | Medium | Update all 3 locations |
| Break existing tests | Low | Low | Update test expectations |

### Testing Plan

1. Navigate to slide with both notes and comments
   - Expected: "Has N comments, has notes, Slide X (Title)"
2. Navigate to slide with only comments
   - Expected: "Has N comments, Slide X (Title)"
3. Navigate to slide with only notes
   - Expected: "has notes, Slide X (Title)"
4. Test in normal mode and slideshow mode
5. Test Comments pane navigation (PageUp/PageDown)

---

## Task 4: Read Actual Notes Instead of "has notes"

### Problem Statement
**Current:** Announces "has notes" prefix
**Required:** Read the actual quick notes content (text between **** markers)

### Current State Analysis

**Where notes are processed:**

| Location | File:Line | Purpose |
|----------|-----------|---------|
| `_has_meeting_notes()` | powerpnt.py:906-923 | Returns True/False if **** markers exist |
| `_get_slide_notes()` | powerpnt.py:855-904 | Gets raw notes text via COM |
| `_clean_notes_text()` | powerpnt.py:925-951 | Extracts text between **** markers |
| `_announce_slide_notes()` | powerpnt.py:953-969 | Called by Ctrl+Alt+N |
| Prefix announcements | Multiple | Currently uses "has notes" string |

**How notes extraction works:**
```python
# _clean_notes_text() lines 937-943
marker_pattern = r'\*{4,}\s*(.*?)\s*\*{4,}'
match = re.search(marker_pattern, notes, re.DOTALL)
if match:
    cleaned = match.group(1)
# Also strips <meeting notes> and <critical notes> tags
```

**Current prefix usage:**
```python
# CustomSlide._get_name() line 1399
if has_notes:
    prefix_parts.append("has notes")  # Static string
```

### Implementation Plan

**Step 1: Add method to get cleaned notes content**

Location: In `PowerPointWorker` class, after `_has_meeting_notes()` (around line 923)

```python
def _get_meeting_notes_content(self):
    """Get the cleaned meeting notes text for current slide.

    v0.0.79: Returns actual notes content instead of just True/False.

    Returns:
        str: Cleaned notes text, or empty string if no meeting notes
    """
    notes = self._get_slide_notes()
    if not notes or '****' not in notes:
        return ""
    return self._clean_notes_text(notes)
```

**Step 2: Cache notes content in worker thread**

Add new cache variable in `__init__` (around line 303):
```python
self._last_notes_content = ""  # v0.0.79: Actual notes text
```

Update `_announce_slide_comments()` to cache content (around line 1028):
```python
# v0.0.79: Cache actual notes content (not just boolean)
notes_content = self._get_meeting_notes_content()
self._last_notes_content = notes_content
self._last_has_notes = bool(notes_content)
```

Update `_cache_slideshow_slide_data()` to cache content (around line 747):
```python
# v0.0.79: Cache actual notes content for slideshow
self._slideshow_notes_content = ""
try:
    notes_page = slide.NotesPage
    placeholder = notes_page.Shapes.Placeholders(2)
    if placeholder.HasTextFrame:
        text_frame = placeholder.TextFrame
        if text_frame.HasText:
            notes_text = text_frame.TextRange.Text.strip()
            if '****' in notes_text:
                self._slideshow_notes_content = self._clean_notes_text(notes_text)
                self._slideshow_has_notes = True
            else:
                self._slideshow_has_notes = False
except Exception as e:
    log.debug(f"Worker: Could not get slideshow notes - {e}")
```

**Step 3: Update CustomSlide._get_name() to use content**

Replace lines 1397-1400:
```python
# v0.0.79: Use actual notes content instead of "has notes"
notes_content = getattr(worker, '_last_notes_content', '')
if notes_content:
    # Truncate if too long for announcement prefix
    if len(notes_content) > 100:
        notes_content = notes_content[:100] + "..."
    prefix_parts.append(notes_content)
```

**Step 4: Update CustomSlideShowWindow._get_name() to use content**

Replace lines 1317-1320:
```python
# v0.0.79: Use actual notes content instead of "has notes"
notes_content = getattr(worker, '_slideshow_notes_content', '')
if notes_content:
    if len(notes_content) > 100:
        notes_content = notes_content[:100] + "..."
    prefix_parts.append(notes_content)
```

**Step 5: Update _announce_slide_comments() to use content**

Replace lines 1044-1045:
```python
# v0.0.79: Use actual notes content instead of "has notes"
if self._last_notes_content:
    notes_display = self._last_notes_content
    if len(notes_display) > 100:
        notes_display = notes_display[:100] + "..."
    prefix_parts.append(notes_display)
```

### Design Considerations

**Length Handling:**
- Quick notes are typically short (1-2 sentences)
- Truncate at 100 characters with "..." if longer
- User can use Ctrl+Alt+N for full content

**Order (considering Task 3):**
After Task 3, order will be: [comments], [notes content], [slide title]

### Risk Assessment

| Risk | Likelihood | Impact | Mitigation |
|------|------------|--------|------------|
| Long notes flood announcement | Medium | Medium | Truncate at 100 chars |
| Notes with special characters | Low | Low | Already cleaned by _clean_notes_text() |
| Empty notes between markers | Low | Low | _clean_notes_text handles this |
| Performance from COM queries | Low | Low | Already cached in worker thread |

**Reference Pitfalls:**
- Pitfall 9: Worker Thread Data Stale at Event Time (using cached values)

### Testing Plan

1. Slide with short notes (under 100 chars)
   - Expected: Full notes content announced
2. Slide with long notes (over 100 chars)
   - Expected: Truncated to 100 chars + "..."
3. Slide with no notes
   - Expected: No notes portion in announcement
4. Slide with notes but no **** markers
   - Expected: No notes portion (not meeting notes)
5. Test Ctrl+Alt+N still reads full content
6. Test in both normal and slideshow modes

---

## Combined Testing Matrix

After all 4 tasks are implemented, run this test matrix:

| Scenario | Normal View | Slideshow | Comments Pane Nav |
|----------|-------------|-----------|-------------------|
| Slide with notes + comments | Test | Test | Test |
| Slide with notes only | Test | Test | Test |
| Slide with comments only | Test | Test | Test |
| Slide with neither | Test | Test | Test |
| Comment from today | Test | N/A | Test |
| Comment from 3 days ago | Test | N/A | Test |
| Comment from 10 days ago | Test | N/A | Test |
| Slide navigation order | Test | Test | Test |

## Implementation Order Recommendation

1. **Task 3 (Announcement Order)** - Simple, low risk, independent
2. **Task 2 (Comment Date)** - Self-contained, no dependencies
3. **Task 4 (Notes Content)** - Depends on order from Task 3
4. **Task 1 (Slide Number)** - Requires investigation first, may be external

## Version Tracking

Each task should increment version:
- Task 3 completion: v0.0.79
- Task 2 completion: v0.0.80
- Task 4 completion: v0.0.81
- Task 1 completion: v0.0.82 (or earlier if quick fix)

---

## Appendix: Key File References

| File | Purpose |
|------|---------|
| `addon/appModules/powerpnt.py` | Main implementation |
| `docs/architecture-decisions.md` | Pattern documentation |
| `docs/history/pitfalls-to-avoid.md` | Known issues |
| `docs/history/failed-approaches.md` | What not to do |
| `docs/reference/nvda-event-timing.md` | Threading model |
| `addon/manifest.ini` | Version declaration |
| `buildVars.py` | Build version |
