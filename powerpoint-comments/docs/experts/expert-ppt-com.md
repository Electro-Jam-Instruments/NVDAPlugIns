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

#### Linear Walkthrough

**Application** is the root object:
- **ActivePresentation** → **Slides** (collection, 1-indexed) → **Slide** (has SlideIndex, Shapes collection, Comments collection)
  - Each **Comment** has: Author, AuthorInitials, Text, DateTime, Replies (Modern Comments)
- **ActiveWindow** → ViewType (int), View → Slide (current slide)

#### 2D Visual Map

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

**CRITICAL:** See `research/PowerPoint-COM-Events-Research.md` for comprehensive details.

### Key Insight: Define Your Own Interface

**DO NOT try to load PowerPoint's type library** - it fails with "Library not registered" in many environments.

**DO define your own EApplication interface class** matching NVDA's pattern. This is reliable and works everywhere.

### Available Events and DISPIDs

| Event | DISPID | When Fired | Notes |
|-------|--------|------------|-------|
| `WindowSelectionChange` | 2001 | Text, shape, or slide selection changes | **RECOMMENDED** - reliable |
| `WindowBeforeRightClick` | 2002 | Before right-click | |
| `WindowBeforeDoubleClick` | 2003 | Before double-click | |
| `PresentationClose` | 2004 | Presentation closing | |
| `PresentationSave` | 2005 | Presentation saved | |
| `PresentationOpen` | 2006 | Presentation opened | |
| `NewPresentation` | 2007 | New presentation created | |
| `PresentationNewSlide` | 2008 | New slide added | |
| `SlideShowBegin` | 2010 | Slideshow starts | |
| `SlideShowEnd` | 2011 | Slideshow ends | |
| `SlideShowNextSlide` | 2013 | Slide advances in slideshow | Used by NVDA |
| `SlideSelectionChanged` | ~2014? | Slide selection changes | DISPID unconfirmed |

### EApplication Interface

**Interface GUID:** `{914934C2-5A91-11CF-8700-00AA0060263B}` (EApplication events)

**Type Library GUID:** `{91493440-5A91-11CF-8700-00AA0060263B}` (PowerPoint type library - DO NOT USE)

### Correct Implementation Pattern

```python
import comtypes
from comtypes.automation import IDispatch
from comtypes import COMObject
import comtypes.client._events
import ctypes

# Step 1: Define the EApplication interface locally
# This is how NVDA does it - define only the events you need
class EApplication(IDispatch):
    """PowerPoint Application Events interface.

    GUID: {914934C2-5A91-11CF-8700-00AA0060263B}
    Define only the events you need with their DISPIDs.
    """
    _iid_ = comtypes.GUID("{914934C2-5A91-11CF-8700-00AA0060263B}")
    _methods_ = []
    _disp_methods_ = [
        # WindowSelectionChange (DISPID 2001) - fires on ANY selection change
        comtypes.DISPMETHOD(
            [comtypes.dispid(2001)],
            None,
            "WindowSelectionChange",
            (["in"], ctypes.POINTER(IDispatch), "sel"),
        ),
        # SlideShowNextSlide (DISPID 2013) - fires during slideshow
        comtypes.DISPMETHOD(
            [comtypes.dispid(2013)],
            None,
            "SlideShowNextSlide",
            (["in"], ctypes.POINTER(IDispatch), "slideShowWindow"),
        ),
    ]


# Step 2: Create event sink that implements the interface
class PowerPointEventSink(COMObject):
    """COM Event Sink for PowerPoint application events."""
    _com_interfaces_ = [EApplication, IDispatch]

    def __init__(self, callback):
        super().__init__()
        self._callback = callback
        self._last_slide = -1

    def WindowSelectionChange(self, sel):
        """Fires on any selection change - reliable for slide detection."""
        if self._callback:
            self._callback("selection_change", sel)

    def SlideShowNextSlide(self, slideShowWindow=None):
        """Fires when slide advances in slideshow mode."""
        if self._callback:
            self._callback("slideshow_slide", slideShowWindow)


# Step 3: Connect using _AdviseConnection (NOT GetEvents)
import comHelper

ppt_app = comHelper.getActiveObject("PowerPoint.Application", dynamic=True)
sink = PowerPointEventSink(my_callback)

# Get IUnknown from sink
sink_iunknown = sink.QueryInterface(comtypes.IUnknown)

# Create advise connection - KEEP THIS REFERENCE!
connection = comtypes.client._events._AdviseConnection(
    ppt_app,
    EApplication,
    sink_iunknown,
)

# Step 4: Pump messages to receive events (see below)
```

### Why NOT to Use GetModule/GetEvents

| Approach | Problem |
|----------|---------|
| `GetModule([GUID, 1, 0])` | "Library not registered" error |
| `GetModule(ppt_app)` | May fail, requires writable comtypes.gen |
| `GetEvents(ppt, sink)` | Requires type library to be loaded first |

**Solution:** Define interface locally + use `_AdviseConnection` directly.

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

NVDA's PowerPoint module (`powerpnt.py`) defines its own `EApplication` class:

```python
# From NVDA source - nvaccess/nvda/source/appModules/powerpnt.py
class EApplication(IDispatch):
    _iid_ = comtypes.GUID("{914934C2-5A91-11CF-8700-00AA0060263B}")
    _methods_ = []
    _disp_methods_ = [
        comtypes.DISPMETHOD([comtypes.dispid(2001)], None, "WindowSelectionChange", ...),
        comtypes.DISPMETHOD([comtypes.dispid(2013)], None, "SlideShowNextSlide", ...),
    ]

class ppEApplicationSink(COMObject):
    _com_interfaces_ = [EApplication, IDispatch]

    def WindowSelectionChange(self, sel):
        # Handle selection change
        ...
```

**Note:** `wireEApplication` does NOT exist in NVDA's source. This was a misunderstanding.

### Threading Considerations

COM events require:
1. **STA Initialization**: `CoInitializeEx(COINIT_APARTMENTTHREADED)`
2. **Message Pump**: Events are delivered via Windows messages
3. **Same Thread**: Event sink must be created on COM thread
4. **Keep Connection Alive**: Store `_AdviseConnection` reference to prevent GC

For NVDA addons:
- Use a dedicated background thread with its own COM initialization
- Or use main thread (already STA) with `core.callLater()` for callbacks

### Event vs Polling Comparison

| Approach | Pros | Cons |
|----------|------|------|
| **Polling** | Simple, works everywhere | CPU usage, latency (poll interval) |
| **Events** | Instant, no CPU waste | Requires interface definition, message pump |

**Recommendation:** Use events with locally-defined interface. This is the proven pattern NVDA uses.

## Performance Notes

1. **Cache COM references** - Don't re-fetch Application repeatedly
2. **Batch reads** - Get all comments at once, not one at a time
3. **Use events over polling** - Instant response, no CPU waste
4. **COM is synchronous** - Long operations block; consider threading
