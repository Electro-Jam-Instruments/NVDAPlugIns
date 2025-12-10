# PowerPoint COM Automation - Expert Knowledge

## Overview

This document contains distilled knowledge for automating Microsoft PowerPoint via COM, specifically for comment navigation.

## Target Environment

| Component | Version |
|-----------|---------|
| PowerPoint | Microsoft 365 (16.0.19426+) |
| Office Automation | COM via comtypes 1.4.11+ |
| NVDA Minimum | 2025.1 |
| NVDA Tested | 2025.3.2 |

**Note:** User's system runs Office 16.0.19426.20170.

## Reference Files

| Topic | File | What It Contains |
|-------|------|------------------|
| Decisions | `decisions.md` | Rationale for technical choices |
| Current Implementation | `MVP_IMPLEMENTATION_PLAN.md` | Phase-by-phase code |
| UIA Focus Management | `windows-accessibility` agent | Comments pane focus |
| COM Library Choice | `nvda-plugins` agent | Why comtypes not pywin32 |
| Research | `research/` folder | Deep analysis documents |

## COM Object Model

### Hierarchy

```
Application
├── ActivePresentation
│   └── Slides (collection, 1-indexed)
│       └── Slide
│           ├── SlideIndex (int)
│           ├── Shapes (collection)
│           └── Comments (collection)
│               └── Comment
│                   ├── Author (string)
│                   ├── AuthorInitials (string)
│                   ├── Text (string)
│                   ├── DateTime (datetime)
│                   └── Replies (collection) [Modern Comments]
└── ActiveWindow
    ├── ViewType (int)
    └── View
        └── Slide (current slide)
```

### Connection

```python
from comtypes.client import GetActiveObject

ppt = GetActiveObject("PowerPoint.Application")
presentation = ppt.ActivePresentation
slide = ppt.ActiveWindow.View.Slide
```

## View Management

### View Type Constants

| Constant | Value | Description |
|----------|-------|-------------|
| ppViewNormal | 9 | Normal editing view |
| ppViewSlideSorter | 5 | Thumbnail grid |
| ppViewNotesPage | 10 | Notes view |
| ppViewOutline | 6 | Outline view |
| ppViewSlideMaster | 3 | Master editing |
| ppViewHandoutMaster | 4 | Handout master |
| ppViewNotesMaster | 5 | Notes master |
| ppViewSlideShow | 1 | Presentation mode |
| ppViewReadingView | 50 | Reading view |

### Get/Set View

```python
# Get current view
current_view = ppt.ActiveWindow.ViewType

# Set to Normal view
ppt.ActiveWindow.ViewType = 9  # ppViewNormal

# Check if Normal
if ppt.ActiveWindow.ViewType == 9:
    # Comments pane accessible
    pass
```

## Comments API

### Accessing Comments

```python
slide = ppt.ActiveWindow.View.Slide
comments = slide.Comments

# Count
count = comments.Count  # 0 if none

# Iterate (1-indexed!)
for i in range(1, comments.Count + 1):
    comment = comments.Item(i)
    print(f"{comment.Author}: {comment.Text}")
```

### Comment Properties

| Property | Type | Description |
|----------|------|-------------|
| Author | string | Display name |
| AuthorInitials | string | Initials |
| Text | string | Comment body (includes @mentions as plain text) |
| DateTime | datetime | When posted |
| Left, Top | int | Position on slide |

### Modern Comments (365)

Modern comments support threading:
```python
comment = slide.Comments.Item(1)
replies = comment.Replies

for i in range(1, replies.Count + 1):
    reply = replies.Item(i)
    print(f"  Reply by {reply.Author}: {reply.Text}")
```

**Note:** Legacy comments (pre-365) don't have Replies collection.

## Opening Comments Pane

### ExecuteMso Command

```python
# Try multiple command names (varies by Office version)
for cmd in ["ReviewShowComments", "ShowComments", "CommentsPane"]:
    try:
        ppt.CommandBars.ExecuteMso(cmd)
        break
    except Exception:
        continue
```

### Toggle vs Show

`ReviewShowComments` toggles the pane. To ensure it's open:
1. Check if pane is visible (via UIA)
2. If not visible, execute command

**Note:** For UIA focus management after opening the pane, see the `windows-accessibility` agent.

## Slide Navigation

### Get Current Slide

```python
slide_index = ppt.ActiveWindow.View.Slide.SlideIndex  # 1-indexed
```

### Navigate to Slide

```python
# By index
ppt.ActiveWindow.View.GotoSlide(3)

# Using Selection
slide = presentation.Slides.Item(3)
slide.Select()
```

### Detect Slide Change

```python
class SlideTracker:
    def __init__(self, ppt):
        self._ppt = ppt
        self._last_index = -1

    def check_changed(self):
        try:
            current = self._ppt.ActiveWindow.View.Slide.SlideIndex
            if current != self._last_index:
                self._last_index = current
                return True
        except Exception:
            pass
        return False
```

## @Mention Parsing

### Pattern

@mentions in PowerPoint are stored as plain text in `Comment.Text`:
```
"@John Smith please review this slide"
```

### Regex Extraction

```python
import re

MENTION_PATTERN = re.compile(
    r'@([A-Za-z\u00C0-\u024F][\w\u00C0-\u024F]*'
    r'(?:[-\'][\w\u00C0-\u024F]+)?'
    r'(?:\s+[A-Za-z\u00C0-\u024F][\w\u00C0-\u024F]*)*)',
    re.UNICODE
)

text = "@John Smith please review"
mentions = MENTION_PATTERN.findall(text)  # ["John Smith"]
```

### User Matching

```python
from difflib import SequenceMatcher

def mentions_user(text, user_name, threshold=0.85):
    mentions = extract_mentions(text)
    for mention in mentions:
        ratio = SequenceMatcher(None,
            mention.lower(),
            user_name.lower()
        ).ratio()
        if ratio >= threshold:
            return True
    return False
```

## Error Handling

### Common Failures

| Scenario | Error | Recovery |
|----------|-------|----------|
| PowerPoint closed | COM error | Reconnect |
| No presentation | AttributeError | Check ActivePresentation |
| No slides | Index error | Check Slides.Count |
| View not Normal | Comments inaccessible | Switch view first |

### Safe COM Wrapper

```python
def safe_com_call(func, *args, fallback=None):
    try:
        return func(*args)
    except Exception:
        return fallback

# Usage
count = safe_com_call(lambda: slide.Comments.Count, fallback=0)
```

## Performance Notes

1. **Cache COM references** - Don't re-fetch Application repeatedly
2. **Batch reads** - Get all comments at once, not one at a time
3. **Avoid polling** - Use events when possible
4. **COM is synchronous** - Long operations block; consider threading
