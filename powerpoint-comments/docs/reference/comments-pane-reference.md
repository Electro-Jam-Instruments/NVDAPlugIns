# Comments Pane Reference

Complete documentation for Comments pane detection, navigation, and comment reformatting.

## UIA Tree Structure

The Comments pane follows this UIA hierarchy:

```
PowerPoint Window (ControlType.Window)
+-- Document (ControlType.Document)
+-- Task Pane Region (ControlType.Pane)
    +-- Comments Pane (ControlType.Pane)
        +-- New Comment Button (ControlType.Button)
        +-- Comments List (ControlType.List) [AutomationId: CommentsList]
            +-- Comment Thread 1 (ControlType.ListItem) [AutomationId: cardRoot_1_GUID]
            │   +-- Author Text (ControlType.Text)
            │   +-- Comment Text (ControlType.Text)
            │   +-- Timestamp Text (ControlType.Text)
            │   +-- More Actions Button (ControlType.Button)
            │   +-- Like Button (ControlType.Button)
            │   +-- Reply Comments (collapsible)
            │       +-- Reply 1 [AutomationId: postRoot_...]
            │       +-- Reply 2
            +-- Comment Thread 2 (ControlType.ListItem) [AutomationId: cardRoot_2_GUID]
            +-- ...
```

### Key Window Classes

| Window Class | Purpose | UIA Behavior |
|-------------|---------|--------------|
| `paneClassDC` | Main slide editing area | UIA disabled by NVDA |
| `mdiClass` | MDI container | UIA disabled by NVDA |
| `NetUIHWND` | Ribbon and task panes | UIA enabled |
| `PodiumParent` | Presentation panel | May require focus for child access |

**Note:** NVDA disables UIA for slide editing windows (`paneClassDC`, `mdiClass`) and uses COM automation instead. The Comments pane uses standard Office task pane patterns with full UIA support.

## UIAutomationId Patterns

These stable identifiers are used to detect Comments pane elements:

| UIAutomationId | Element Type | Notes |
|----------------|--------------|-------|
| `NewCommentButton` | Button | First element when F6 to pane |
| `CommentsList` | List container | Parent of all comment threads |
| `cardRoot_*` | Comment thread | Prefix for thread containers |
| `postRoot_*` | Reply comment | Prefix for reply cards |
| `firstPaneElement*` | Pane container | Prefix for pane boundary |

## Detecting Comments Pane Focus

```python
def _is_in_comments_pane(self, obj):
    """Check if object is in Comments pane."""
    uia_id = getattr(obj, 'UIAAutomationId', '') or ''

    # Direct matches
    if uia_id in ('NewCommentButton', 'CommentsList'):
        return True

    # Prefix matches
    if uia_id.startswith('cardRoot_'):  # Comment threads
        return True
    if uia_id.startswith('postRoot_'):  # Replies
        return True
    if uia_id.startswith('firstPaneElement'):  # Pane container
        return True

    return False
```

## Comment Name Formats from PowerPoint

PowerPoint provides comment names in three formats. Understanding these is critical for reformatting.

### Thread Comments (cardRoot_)

Format: `"[Resolved ]Comment thread started by Author, with N replies"`

Examples:
- `"Comment thread started by John Smith, with 2 replies"`
- `"Resolved Comment thread started by Jane Doe, with 0 replies"`

### Reply Comments (postRoot_)

Format: `"Comment by Author on Month Day, Year, Time"`

Example:
- `"Comment by John Smith on January 5, 2026, 2:30 PM"`

### Task Status Updates

Format: `"Task updated by Author on Month Day, Year, Time"`

Example:
- `"Task updated by Jane Doe on January 6, 2026, 10:15 AM"`

## Comment Reformatting Algorithm

```python
def _reformat_comment_name(self, name):
    """Reformat PowerPoint's verbose comment name to 'Author: text'.

    Returns reformatted name or None if not a comment.
    """
    if not name:
        return None

    # CRITICAL: Normalize whitespace - PowerPoint uses U+00A0 (non-breaking space)
    name_normalized = re.sub(r'\s+', ' ', name)

    # Check for resolved status
    is_resolved = name_normalized.startswith("Resolved ")
    if is_resolved:
        name_normalized = name_normalized[9:]  # Remove "Resolved "

    author = None

    # Pattern 1: Thread comments - "Comment thread started by Author, with N replies"
    if " started by " in name_normalized:
        author_part = name_normalized.split(" started by ", 1)[1]
        # Author ends at ", with"
        if ", with" in author_part:
            author = author_part.split(", with")[0].strip()

    # Pattern 2: Reply comments - "Comment by Author on Month Day, Year"
    elif name_normalized.startswith("Comment by "):
        after_prefix = name_normalized[11:]  # Skip "Comment by "
        # Author ends at " on " (date)
        if " on " in after_prefix:
            author = after_prefix.split(" on ")[0].strip()

    # Pattern 3: Task status - "Task updated by Author on Month Day, Year"
    elif name_normalized.startswith("Task updated by "):
        after_prefix = name_normalized[16:]  # Skip "Task updated by "
        if " on " in after_prefix:
            author = after_prefix.split(" on ")[0].strip()

    if not author:
        return None

    # Build reformatted name
    prefix = "Resolved - " if is_resolved else ""
    return f"{prefix}{author}:"
```

## Non-Breaking Space Pitfall

**CRITICAL:** PowerPoint uses U+00A0 (non-breaking space) in comment names, not regular spaces (U+0020).

```python
# WRONG - Fails to match because of non-breaking spaces
if "started by" in comment_name:  # May fail!

# CORRECT - Normalize all whitespace first
name_normalized = re.sub(r'\s+', ' ', comment_name)
if "started by" in name_normalized:  # Works!
```

Without this normalization, string operations like `.split()` and `.startswith()` will fail silently.

## Auto-Tab Behavior

When user presses F6 to enter Comments pane, focus lands on NewCommentButton. We auto-tab to first comment for better UX:

```python
def event_gainFocus(self, obj, nextHandler):
    uia_id = getattr(obj, 'UIAAutomationId', '') or ''

    # Auto-tab from NewCommentButton on initial pane entry
    if uia_id == 'NewCommentButton':
        if not getattr(self, '_in_comments_pane', False):
            self._in_comments_pane = True
            log.info("Entering Comments pane - auto-tabbing to first comment")
            # Send synthetic Tab keypress
            from keyboardHandler import KeyboardInputGesture
            KeyboardInputGesture.fromName("tab").send()
            return  # Don't announce the button

    nextHandler()
```

**Flags used:**
- `_in_comments_pane` - True when focus is in Comments pane
- `_pending_auto_focus` - Set before navigation to enable auto-tab on landing

## PageUp/PageDown Navigation from Comments Pane

When user presses PageUp/PageDown while in Comments pane, we navigate slides:

```python
def script_navigateSlide(self, gesture):
    """Handle PageUp/PageDown in Comments pane."""
    direction = 1 if gesture.mainKeyName == "pageDown" else -1

    # Check if in Comments pane
    focus = api.getFocusObject()
    if self._is_in_comments_pane(focus):
        # Request slide navigation from worker thread
        self._worker.request_navigate(direction, from_comments_pane=True)
        # Don't pass through - we handle it
        return

    # Not in Comments pane - pass through to PowerPoint
    gesture.send()
```

## Checking if Comments Pane is Open

Use `GetPressedMso()` to check pane visibility before toggling:

```python
def _is_comments_pane_visible(self):
    """Check if Comments pane is currently open."""
    # Try multiple command names (varies by Office version)
    for cmd in ["CommentsPane", "ReviewShowComments", "ShowComments"]:
        try:
            state = self._ppt_app.CommandBars.GetPressedMso(cmd)
            if state:  # True or -1 means pressed/active
                return True
        except:
            continue
    return False
```

**Why this matters:** `ExecuteMso("ReviewShowComments")` is a toggle. Without checking state first, you might close the pane when trying to open it.

## Opening Comments Pane

```python
def _open_comments_pane(self):
    """Open Comments pane if not already open."""
    if self._is_comments_pane_visible():
        return  # Already open

    # Try multiple command names
    for cmd in ["ReviewShowComments", "ShowComments", "CommentsPane"]:
        try:
            self._ppt_app.CommandBars.ExecuteMso(cmd)
            return
        except:
            continue
```

## Speech Cancellation Logic

When reformatting comments, we cancel NVDA's default speech and announce our reformatted version:

```python
def event_gainFocus(self, obj, nextHandler):
    # ... comment detection ...

    reformatted = self._reformat_comment_name(obj.name)
    if reformatted:
        # Cancel NVDA's default announcement
        # BUT: Skip if we just navigated (don't cut off slide title)
        if not getattr(self, '_just_navigated', False):
            speech.cancelSpeech()
        else:
            self._just_navigated = False

        # Announce reformatted name
        ui.message(reformatted)
        return

    nextHandler()
```

**The `_just_navigated` flag** prevents cutting off slide title announcements when PageUp/PageDown triggers a slide change followed by focus change.

## Complete Comment Detection Flow

```
User presses Tab in Comments pane
         │
         ▼
event_gainFocus fires
         │
         ▼
Get UIAAutomationId ─────────────────────────────┐
         │                                        │
         ▼                                        │
Is it cardRoot_* or postRoot_*?                   │
         │                                        │
    ┌────┴────┐                                   │
    │ YES     │ NO                                │
    ▼         ▼                                   │
Get obj.name  Pass to nextHandler()               │
    │                                             │
    ▼                                             │
Normalize whitespace (U+00A0 → U+0020)            │
    │                                             │
    ▼                                             │
Parse "started by" / "Comment by" / "Task updated"│
    │                                             │
    ▼                                             │
Extract author, detect resolved status            │
    │                                             │
    ▼                                             │
Cancel speech (if not _just_navigated)            │
    │                                             │
    ▼                                             │
Announce: "Author:" or "Resolved - Author:"       │
```
