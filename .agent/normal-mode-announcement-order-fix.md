# Normal Mode: Fix Announcement Order (has notes, has N comments BEFORE slide)

## Current Problem

**Current Order (WRONG):**
1. NVDA built-in: "Slide 1, Title"
2. Our addon: "Has 1 comment"
3. Our addon: "has notes"

**Desired Order:**
1. Our addon: "has notes"
2. Our addon: "Has 1 comment"
3. NVDA built-in: "Slide 1, Title"

## Root Cause

### Current Flow
```
User changes slide (arrow key, click, etc.)
  ↓
PowerPoint fires WindowSelectionChange event
  ↓
Worker thread: on_slide_changed_event()
  ↓
Worker thread: _announce_slide_comments()  ← Announces comments/notes
  ↓
MEANWHILE (in parallel or slightly before):
  ↓
NVDA built-in: Detects focus change
  ↓
NVDA built-in: Announces "Slide X, Title"
```

**Problem:** Our worker thread announcements happen AFTER (or in race with) NVDA's built-in slide announcement.

## Solution Approach

### Option 1: Intercept Focus Event (RECOMMENDED)

Override `event_gainFocus` for slide objects to announce comments/notes BEFORE calling `nextHandler()`.

**Pattern:**
```python
def event_gainFocus(self, obj, nextHandler):
    # Check if obj is a slide
    if self._is_slide_object(obj):
        # Announce comments/notes FIRST
        if self._worker:
            slide_index = self._worker._get_current_slide_index()
            comments_count = len(self._worker._get_comments_on_current_slide())
            has_notes = self._worker._has_meeting_notes()

            # Build announcement
            parts = []
            if has_notes:
                parts.append("has notes")
            if comments_count > 0:
                parts.append(f"Has {comments_count} comment{'s' if comments_count != 1 else ''}")

            if parts:
                import ui
                ui.message(", ".join(parts))

        # THEN let NVDA announce slide normally
        nextHandler()
    else:
        # Not a slide, normal handling
        nextHandler()
```

**Advantages:**
- Clean interception point
- Guaranteed order (our announcement → NVDA's announcement)
- No timing dependencies

**Challenges:**
- Need to identify slide objects reliably
- Already using `event_gainFocus` for comment card reformatting
- Must not break existing comment navigation logic

### Option 2: Use `getSpeechTextForProperties` Override

Override how NVDA generates speech for slide objects.

**Pattern:**
```python
def getSpeechTextForProperties(self, obj, reason=controlTypes.OutputReason.QUERY, *args, **kwargs):
    text_gen = super().getSpeechTextForProperties(obj, reason, *args, **kwargs)

    if self._is_slide_object(obj):
        # Prepend comments/notes to speech
        if self._worker:
            has_notes = self._worker._has_meeting_notes()
            comments_count = len(self._worker._get_comments_on_current_slide())

            if has_notes:
                yield "has notes"
            if comments_count > 0:
                yield f"Has {comments_count} comment{'s' if comments_count != 1 else ''}"

    # Then yield normal speech
    for item in text_gen:
        yield item
```

**Advantages:**
- Modifies speech generation directly
- Integrated with NVDA's speech system

**Challenges:**
- More complex API
- May interact with verbosity settings unexpectedly
- Harder to debug

### Option 3: Cancel and Replace (Current approach for comments)

Cancel NVDA's announcement and replace with our own.

**Already used for comment cards (lines 1125-1166):**
```python
def event_gainFocus(self, obj, nextHandler):
    # For comment cards
    if obj.UIAElement.CurrentClassName == "CommentCard":
        cancelSpeech()
        # ... reformat and announce
        return  # Don't call nextHandler - we replaced the announcement
```

**For slides:**
```python
def event_gainFocus(self, obj, nextHandler):
    if self._is_slide_object(obj):
        cancelSpeech()

        # Announce OUR version with comments/notes first
        parts = []
        if has_notes:
            parts.append("has notes")
        if comments_count > 0:
            parts.append(f"Has {comments_count} comment{'s' if comments_count != 1 else ''}")

        # Then get slide info from NVDA
        slide_speech = super().getSpeechTextForProperties(obj, ...)
        parts.extend(slide_speech)

        ui.message(", ".join(parts))
        return  # Don't call nextHandler
```

**Advantages:**
- Full control over announcement
- Guaranteed order

**Challenges:**
- Must reproduce ALL of NVDA's slide announcement logic
- Risk of missing verbosity settings, etc.
- More maintenance burden

## Recommended Implementation: Option 1 Modified

**Strategy:** Use `event_gainFocus` but DON'T cancel speech. Just prepend our announcement.

```python
def event_gainFocus(self, obj, nextHandler):
    try:
        # Existing comment card logic (lines 1125-1166)
        if hasattr(obj, 'UIAElement') and obj.UIAElement:
            if obj.UIAElement.CurrentClassName == "CommentCard":
                # ... existing comment card reformatting ...
                return

        # NEW: Slide announcement prepending
        if self._is_slide_object(obj):
            if self._worker:
                slide_index = self._worker._get_current_slide_index()
                has_notes = self._worker._has_meeting_notes()
                comments = self._worker._get_comments_on_current_slide()
                comments_count = len(comments)

                # Build prepended announcement
                parts = []
                if has_notes:
                    parts.append("has notes")
                if comments_count > 0:
                    parts.append(f"Has {comments_count} comment{'s' if comments_count != 1 else ''}")

                if parts:
                    import ui
                    ui.message(", ".join(parts))
                    # Small delay to ensure our announcement finishes first
                    import time
                    time.sleep(0.1)

                # Open comments pane if needed
                if comments_count > 0:
                    self._worker._open_comments_pane()

        # Always call nextHandler for normal NVDA processing
        nextHandler()

    except Exception as e:
        log.error(f"event_gainFocus error: {e}", exc_info=True)
        nextHandler()
```

### Identifying Slide Objects

**Need to determine:** What properties identify a slide object?

Possible checks:
- `obj.role == controlTypes.Role.PANE` and in slide area?
- `obj.windowClassName` specific to slides?
- Check parent chain for slide container?

**Action:** Add logging to discover slide object properties during normal slide navigation.

## Implementation Plan

### Phase 1: Discovery (v0.0.68)
1. Add diagnostic logging in `event_gainFocus` to capture obj properties during slide navigation
2. Log: `role`, `name`, `windowClassName`, `UIAElement.CurrentClassName`
3. User tests: Navigate between slides in normal mode
4. Review logs to identify slide object signature

### Phase 2: Implementation (v0.0.69)
1. Implement `_is_slide_object()` helper based on Phase 1 findings
2. Add slide announcement prepending in `event_gainFocus`
3. Test announcement order in normal mode

### Phase 3: Refinement (v0.0.70)
1. Remove `time.sleep()` if not needed
2. Ensure no conflicts with existing comment card logic
3. Clean up logs

## Testing Checklist

- [ ] Normal mode: Navigate slides with arrow keys
  - Hear: "has notes, Has 1 comment, Slide 1, Title"
- [ ] Normal mode: Navigate slides without comments/notes
  - Hear: "Slide 1, Title" (no extra announcements)
- [ ] Normal mode: Navigate from Comments pane
  - Existing logic should still work
- [ ] Comment card navigation still works
  - Date/time reformatting preserved
- [ ] Auto-tab from NewCommentButton still works

## File Changes
`powerpoint-comments/addon/appModules/powerpnt.py`
- Lines 1125-1200: `event_gainFocus` method (existing comment card logic)
- Add: `_is_slide_object()` helper method
- Add: Slide announcement prepending logic
