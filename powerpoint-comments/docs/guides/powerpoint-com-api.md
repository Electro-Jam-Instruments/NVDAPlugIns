# PowerPoint COM API Guide

How to automate Microsoft PowerPoint via COM for the NVDA addon.

## Connecting to PowerPoint

**CRITICAL: Use `comHelper.getActiveObject()` NOT direct `GetActiveObject()`**

NVDA runs with UIAccess privileges which prevents direct COM access.

```python
# CORRECT - Use NVDA's comHelper
import comHelper
ppt_app = comHelper.getActiveObject("PowerPoint.Application", dynamic=True)

# WRONG - Fails with UIAccess privilege error
from comtypes.client import GetActiveObject
ppt_app = GetActiveObject("PowerPoint.Application")  # WinError -2147221021
```

## COM Object Hierarchy

```
Application
├── ActivePresentation
│   └── Slides (collection, 1-indexed)
│       └── Slide
│           ├── SlideIndex (int)
│           ├── Shapes (collection)
│           ├── Comments (collection)
│           └── NotesPage.Shapes[2].TextFrame.TextRange.Text
└── ActiveWindow
    ├── ViewType (int)
    └── View
        └── Slide (current slide)
```

## Getting Current Slide

```python
slide = ppt_app.ActiveWindow.View.Slide
slide_index = slide.SlideIndex  # 1-indexed
```

## Accessing Comments

```python
comments = slide.Comments

# Count (0 if none)
count = comments.Count

# Iterate (1-indexed!)
for i in range(1, comments.Count + 1):
    comment = comments.Item(i)
    author = comment.Author
    text = comment.Text
```

## Accessing Slide Notes

```python
notes_text = ""
try:
    notes_page = slide.NotesPage
    # Notes text is in placeholder 2 (not Item(2)!)
    placeholder = notes_page.Shapes.Placeholders(2)
    if placeholder.HasTextFrame:
        text_frame = placeholder.TextFrame
        if text_frame.HasText:
            notes_text = text_frame.TextRange.Text.strip()
except:
    pass
```

### Extracting Meeting Notes (with **** markers)

Our addon treats text between `****` markers as "meeting notes":

```python
import re

def has_meeting_notes(notes_text):
    """Check if notes contain **** markers."""
    return '****' in notes_text

def extract_meeting_notes(notes_text):
    """Extract only text between **** markers."""
    marker_pattern = r'\*{4,}\s*(.*?)\s*\*{4,}'
    match = re.search(marker_pattern, notes_text, re.DOTALL)
    if match:
        content = match.group(1).strip()
        # Remove <meeting notes> and <critical notes> tags
        content = re.sub(r'</?meeting\s*notes>', '', content, flags=re.IGNORECASE)
        content = re.sub(r'</?critical\s*notes>', '', content, flags=re.IGNORECASE)
        return content.strip()
    return ""
```

## COM Events - The Right Way

**DO NOT try to load PowerPoint's type library** - it fails with "Library not registered".

**DO define the EApplication interface locally:**

```python
import comtypes
from comtypes.automation import IDispatch
from comtypes import COMObject, GUID
from comtypes.client._events import _AdviseConnection
import ctypes

class EApplication(IDispatch):
    """PowerPoint Application Events - defined locally."""
    _iid_ = GUID("{914934C2-5A91-11CF-8700-00AA0060263B}")
    _methods_ = []
    _disp_methods_ = [
        comtypes.DISPMETHOD(
            [comtypes.dispid(2001)], None, "WindowSelectionChange",
            (["in"], ctypes.POINTER(IDispatch), "sel"),
        ),
        comtypes.DISPMETHOD(
            [comtypes.dispid(2010)], None, "SlideShowBegin",
            (["in"], ctypes.POINTER(IDispatch), "wn"),
        ),
        comtypes.DISPMETHOD(
            [comtypes.dispid(2012)], None, "SlideShowEnd",
            (["in"], ctypes.POINTER(IDispatch), "pres"),
        ),
        comtypes.DISPMETHOD(
            [comtypes.dispid(2013)], None, "SlideShowNextSlide",
            (["in"], ctypes.POINTER(IDispatch), "slideShowWindow"),
        ),
    ]

class PowerPointEventSink(COMObject):
    _com_interfaces_ = [EApplication, IDispatch]

    def WindowSelectionChange(self, sel):
        # Fires on slide/shape/text selection change in edit mode
        pass

    def SlideShowNextSlide(self, slideShowWindow):
        # Fires when slide advances in slideshow
        pass
```

## Connecting the Event Sink

```python
sink = PowerPointEventSink()
sink_iunknown = sink.QueryInterface(comtypes.IUnknown)
connection = _AdviseConnection(ppt_app, EApplication, sink_iunknown)
# KEEP connection reference alive!
```

## Key Event DISPIDs

| Event | DISPID | When Fired |
|-------|--------|------------|
| WindowSelectionChange | 2001 | Edit mode - selection changes |
| SlideShowBegin | 2010 | Slideshow starts |
| SlideShowEnd | 2012 | Slideshow ends |
| SlideShowNextSlide | 2013 | Slideshow - slide advances |

## Message Pump Requirement

COM events require Windows message processing:

```python
from ctypes import windll, byref
from ctypes.wintypes import MSG

def pump_messages(timeout_ms=500):
    user32 = windll.user32
    user32.MsgWaitForMultipleObjects(0, None, False, timeout_ms, 0x04FF)

    msg = MSG()
    while user32.PeekMessageW(byref(msg), None, 0, 0, 0x0001):
        user32.TranslateMessage(byref(msg))
        user32.DispatchMessageW(byref(msg))
```

## View Type Constants

| Constant | Value | Description |
|----------|-------|-------------|
| ppViewNormal | 9 | Normal editing view |
| ppViewSlideShow | 1 | Presentation mode |
| ppViewSlideSorter | 5 | Thumbnail grid |

## Error Handling

```python
def safe_com_call(func, fallback=None):
    try:
        return func()
    except Exception:
        return fallback

count = safe_com_call(lambda: slide.Comments.Count, fallback=0)
```

## Detecting Modes

### Check if in Slideshow Mode

```python
def is_in_slideshow(ppt_app):
    """Returns True if a slideshow is running."""
    try:
        return ppt_app.SlideShowWindows.Count > 0
    except:
        return False
```

### Check Current View Type

```python
PP_VIEW_NORMAL = 9
PP_VIEW_SLIDESHOW = 1
PP_VIEW_SLIDE_SORTER = 5

def get_view_type(ppt_app):
    """Returns current view type constant."""
    try:
        return ppt_app.ActiveWindow.ViewType
    except:
        return None

# Usage
if get_view_type(ppt_app) == PP_VIEW_NORMAL:
    # Edit mode
    pass
```

### Check if Presentation is Open

```python
def has_active_presentation(ppt_app):
    """Returns True if a presentation is open."""
    try:
        return (ppt_app.Presentations.Count > 0 and
                ppt_app.ActiveWindow is not None)
    except:
        return False
```

### Check if PowerPoint is Running

```python
def get_powerpoint():
    """Returns PowerPoint app or None if not running."""
    try:
        return comHelper.getActiveObject("PowerPoint.Application", dynamic=True)
    except OSError:
        return None  # PowerPoint not running
```

## What `dynamic=True` Does

The `dynamic=True` parameter tells comtypes to use late binding (IDispatch) instead of early binding. This:
- Avoids needing type library
- Works even when type library registration fails
- Slower but more reliable

Always use `dynamic=True` for PowerPoint in NVDA addons.

## Checking Ribbon/Pane State

Use `GetPressedMso()` to check if a ribbon button is pressed or pane is open:

```python
def is_comments_pane_visible(ppt_app):
    """Check if Comments pane is open."""
    # Try multiple command names (varies by Office version)
    for cmd in ["CommentsPane", "ReviewShowComments", "ShowComments"]:
        try:
            state = ppt_app.CommandBars.GetPressedMso(cmd)
            if state:  # True or -1 means pressed/active
                return True
        except:
            continue
    return False
```

**Why this matters:** `ExecuteMso()` toggles state. Without checking first, you might close a pane when trying to open it.

### Executing Ribbon Commands

```python
def open_comments_pane(ppt_app):
    """Open Comments pane if not already open."""
    if is_comments_pane_visible(ppt_app):
        return  # Already open

    for cmd in ["ReviewShowComments", "ShowComments", "CommentsPane"]:
        try:
            ppt_app.CommandBars.ExecuteMso(cmd)
            return
        except:
            continue
```

## Multiple Presentations - sel.Parent Pattern

When handling `WindowSelectionChange` with multiple presentations open, use `sel.Parent` to get the correct window:

```python
def WindowSelectionChange(self, sel):
    # WRONG - May return wrong window if multiple presentations open
    window = self._ppt_app.ActiveWindow

    # CORRECT - Get specific window from selection
    try:
        window = sel.Parent  # DocumentWindow that owns the selection
    except:
        window = self._ppt_app.ActiveWindow  # Fallback
```

## Important Notes

1. **COM collections are 1-indexed** - Use `range(1, count + 1)`
2. **Cache COM references** - Don't re-fetch Application repeatedly
3. **COM is synchronous** - Long operations block; use threading
4. **STA required** - Use `CoInitializeEx(COINIT_APARTMENTTHREADED)` on worker threads
5. **Keep event connection alive** - Store `_AdviseConnection` in instance variable
6. **Use sel.Parent** - For correct window with multiple presentations
