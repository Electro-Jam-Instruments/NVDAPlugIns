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

## COM Events (Event-Driven Approach)

PowerPoint exposes application-level events via COM. This is superior to polling for detecting slide changes.

### Available Events

| Event | When Fired | Use For |
|-------|------------|---------|
| `SlideSelectionChanged` | Slide thumbnail selection changes | Detecting slide navigation |
| `WindowSelectionChange` | Text, shape, or slide selection changes | Broader selection tracking |
| `SlideShowBegin` | Slideshow starts | Disable monitoring |
| `SlideShowEnd` | Slideshow ends | Re-enable monitoring |
| `PresentationClose` | File closed | Cleanup |

### EApplication Interface

PowerPoint events use the `EApplication` interface (ProgID: `PowerPoint.Application`).

**PowerPoint Type Library GUID:** `{91493440-5A91-11CF-8700-00AA0060263B}`

```python
from comtypes import COMObject
from comtypes.client import GetEvents, GetModule
import comHelper

# Step 1: Load PowerPoint type library to get EApplication interface
ppt_gen = GetModule(['{91493440-5A91-11CF-8700-00AA0060263B}', 1, 0])
EApplication = ppt_gen.EApplication  # The events interface

class PowerPointEventSink(COMObject):
    """Receives PowerPoint application events."""

    # MUST set this AFTER loading type library
    _com_interfaces_ = [EApplication]

    def SlideSelectionChanged(self, SldRange):
        """Called when slide selection changes in thumbnail pane."""
        # SldRange is a SlideRange object
        if SldRange and SldRange.Count > 0:
            slide_index = SldRange.Item(1).SlideIndex
            # Handle slide change

    def WindowSelectionChange(self, Sel):
        """Called when selection changes in window."""
        # Sel is a Selection object with Type property:
        # 0=ppSelectionNone, 1=ppSelectionSlides, 2=ppSelectionShapes, 3=ppSelectionText
        pass

# Step 2: Connect to PowerPoint (use comHelper for NVDA!)
ppt = comHelper.getActiveObject("PowerPoint.Application", dynamic=True)

# Step 3: Create sink and connect
sink = PowerPointEventSink()
connection = GetEvents(ppt, sink)

# Step 4: CRITICAL - Pump Windows messages for events to fire
# See "Message Pump Requirement" section below
```

### Message Pump Requirement

**COM events are delivered via Windows messages.** Without a message pump, events will NOT fire.

```python
from ctypes import windll, byref
from ctypes.wintypes import MSG

def pump_messages(timeout_ms=500):
    """Process Windows messages to receive COM events."""
    user32 = windll.user32
    QS_ALLINPUT = 0x04FF
    PM_REMOVE = 0x0001

    # Wait for messages with timeout
    user32.MsgWaitForMultipleObjects(0, None, False, timeout_ms, QS_ALLINPUT)

    # Process all pending messages
    msg = MSG()
    while user32.PeekMessageW(byref(msg), None, 0, 0, PM_REMOVE):
        user32.TranslateMessage(byref(msg))
        user32.DispatchMessageW(byref(msg))

# In your main loop:
while not stop_event.is_set():
    pump_messages(timeout_ms=500)
```

### NVDA's Built-in Pattern

NVDA's PowerPoint module uses `ppEApplicationSink` for event handling:

```python
# From nvdaBuiltin.appModules.powerpnt
class ppEApplicationSink(COMObject):
    _com_interfaces_ = [wireEApplication]

    def wireWindowSelectionChange(self, pSel):
        """Handle selection change in PowerPoint window."""
        # NVDA processes selection and updates focus
```

### Threading Considerations

COM events require:
1. **STA Initialization**: `CoInitializeEx(COINIT_APARTMENTTHREADED)`
2. **Message Pump**: Events are delivered via Windows messages
3. **Same Thread**: Event sink must be created on COM thread

For NVDA addons, the pattern is:
- Create event sink on main thread (already STA)
- Or use `wx.CallAfter()` / `core.callLater()` for thread safety

### Event vs Polling Comparison

| Approach | Pros | Cons |
|----------|------|------|
| **Polling** | Simple, works everywhere | CPU usage, latency (poll interval) |
| **Events** | Instant, no CPU waste | Requires COM event setup, message pump |

**Recommendation:** Use events when possible; fall back to polling if events fail.

## Performance Notes

1. **Cache COM references** - Don't re-fetch Application repeatedly
2. **Batch reads** - Get all comments at once, not one at a time
3. **Use events over polling** - Instant response, no CPU waste
4. **COM is synchronous** - Long operations block; consider threading
