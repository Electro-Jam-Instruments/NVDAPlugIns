# PowerPoint UI Automation (UIA) Research Documentation

## Executive Summary

This document provides comprehensive research on Microsoft PowerPoint's UI Automation implementation, accessibility tree structure, supported patterns, and known limitations. This research is intended to support development of an NVDA plug-in for PowerPoint accessibility.

**Key Findings:**
- PowerPoint uses a hybrid approach: UIA for modern UI elements, COM automation for document content access
- Microsoft provides custom UIA properties for PowerPoint (ViewType with 30 view states)
- NVDA's PowerPoint appModule disables UIA for certain window classes and relies on COM automation instead
- Significant limitations exist with shape/object accessibility in slideshow mode
- Comments pane accessibility follows standard task pane patterns with specific keyboard navigation

---

## Table of Contents

1. [UIA Tree Structure and Hierarchy](#1-uia-tree-structure-and-hierarchy)
2. [UIA Control Types Used by PowerPoint](#2-uia-control-types-used-by-powerpoint)
3. [UIA Patterns Supported by PowerPoint](#3-uia-patterns-supported-by-powerpoint)
4. [PowerPoint-Specific UIA Properties](#4-powerpoint-specific-uia-properties)
5. [Comments Pane Accessibility](#5-comments-pane-accessibility)
6. [Known Limitations and Workarounds](#6-known-limitations-and-workarounds)
7. [Exploration Tools and Techniques](#7-exploration-tools-and-techniques)
8. [Code Examples](#8-code-examples)
9. [NVDA PowerPoint AppModule Analysis](#9-nvda-powerpoint-appmodule-analysis)
10. [References](#10-references)

---

## 1. UIA Tree Structure and Hierarchy

### Root Window Structure

PowerPoint's UIA tree follows this general hierarchy:

#### Linear Walkthrough

1. **Desktop (Root)** - Top level
2. **PowerPoint Application Window (POWERPNT.EXE)** contains:
   - **Ribbon (NetUIHWND)** - Contains Ribbon Tabs (Tab control), Ribbon Groups (Pane elements), Quick Access Toolbar
   - **Document Area (mdiClass / paneClassDC)** - Contains Slide Editing Pane (ViewSlide), Thumbnail Pane (ViewThumbnails), Notes Pane (ViewNotesText), Outline Pane (ViewOutline)
   - **Task Panes** - Comments, Accessibility Checker, etc. The Comments Pane contains Comment Threads (List-like structure) and Individual Comments
   - **Status Bar**

#### 2D Visual Map

```
Desktop (Root)
└── PowerPoint Application Window (POWERPNT.EXE)
    ├── Ribbon (NetUIHWND)
    │   ├── Ribbon Tabs (Tab control)
    │   ├── Ribbon Groups (Pane elements)
    │   └── Quick Access Toolbar
    ├── Document Area (mdiClass / paneClassDC)
    │   ├── Slide Editing Pane (Pane - ViewSlide)
    │   ├── Thumbnail Pane (Pane - ViewThumbnails)
    │   ├── Notes Pane (Pane - ViewNotesText)
    │   └── Outline Pane (Pane - ViewOutline)
    ├── Task Panes (Comments, Accessibility Checker, etc.)
    │   └── Comments Pane (Pane)
    │       ├── Comment Threads (List-like structure)
    │       └── Individual Comments
    └── Status Bar
```

### Critical Window Classes

PowerPoint uses specific window classes that affect UIA behavior:

| Window Class | Purpose | UIA Behavior |
|-------------|---------|--------------|
| `paneClassDC` | Main slide editing area | UIA disabled by NVDA (uses COM instead) |
| `mdiClass` | MDI container for documents | UIA disabled by NVDA (generic container) |
| `screenClass` | Slideshow presentation mode | UIA disabled by NVDA |
| `NetUIHWND` | Ribbon and modern UI elements | UIA enabled |
| `PodiumParent` | Presentation panel container | May require focus for child access |

### PodiumParent Behavior

The `PodiumParent` panel uses deferred UIA provider loading:
- Child elements may not appear until window receives focus
- Using "Watch Focus" in Inspect.exe helps reveal elements
- Provider information shows: `[providerId:0x0 Main(parent link):Unidentified Provider (unmanaged:mso.dll)]`

---

## 2. UIA Control Types Used by PowerPoint

### Common Control Types

| Control Type | PowerPoint Usage |
|-------------|------------------|
| `Pane` | Main content areas (slides, notes, thumbnails, task panes) |
| `Document` | Not typically used (content accessed via COM) |
| `List` | Slide thumbnails list, comment lists |
| `ListItem` | Individual slides in thumbnail view |
| `Button` | Ribbon buttons, task pane actions |
| `Tab` | Ribbon tabs |
| `TabItem` | Individual ribbon tab |
| `Text` | Text labels |
| `Edit` | Text input fields (notes, comments) |
| `Custom` | Some specialized PowerPoint controls |
| `Image` | Pictures and media on slides |

### Shape Type to NVDA Role Mapping

NVDA maps PowerPoint shape types to accessibility roles:

| PowerPoint Shape Type | NVDA Role |
|----------------------|-----------|
| Chart | Chart |
| Group | GroupBox |
| Embedded OLE Object | EmbeddedObject |
| Line | Line |
| Picture | Picture |
| Text Box | TextBox |
| Table | Table |
| Diagram/SmartArt | Diagram |
| Media (Audio) | Audio |
| Media (Video) | Video |
| Action Button | Shape with action name exposed |

---

## 3. UIA Patterns Supported by PowerPoint

### Pattern Support Matrix

| UIA Pattern | Support Level | Notes |
|------------|---------------|-------|
| TextPattern | Limited | Not standard; NVDA uses COM TextRange |
| SelectionPattern | Partial | Works for ribbon, thumbnails |
| ScrollPattern | Yes | Task panes, lists with scrollbars |
| ValuePattern | Partial | Some edit controls |
| InvokePattern | Yes | Buttons, menu items |
| ExpandCollapsePattern | Yes | Ribbon groups, comment threads |
| SelectionItemPattern | Yes | Ribbon tabs, thumbnails |
| GridPattern | Limited | Tables may expose |
| TablePattern | Limited | Tables in slides |
| TextRangePattern | Not implemented | Use COM TextRange instead |

### Why PowerPoint Does Not Fully Implement UIA TextPattern

PowerPoint's slide content is not exposed through standard UIA TextPattern. Microsoft's implementation provides:
- Custom UIA properties for view identification
- COM automation through the PowerPoint object model
- Direct IDispatch access to shapes, text frames, and ranges

---

## 4. PowerPoint-Specific UIA Properties

### Custom Properties

Microsoft exposes custom UIA properties specifically for PowerPoint (available in Microsoft 365 Version 2012, Build 13530+):

#### ViewType Property

- **GUID:** `{F065BAA7-2794-48B6-A927-193DA1540B84}`
- **Applies To:** Pane Control Type
- **Variant Type:** VT_I4 (integer)
- **Default Value:** 30 (ViewUnknown)

#### PowerPointViewType Enumeration (30 Values)

| Value | Enum Name | Description |
|-------|-----------|-------------|
| 1 | ViewSlide | The element is the Slide Pane |
| 2 | ViewSlideMaster | The element is the Slide Master Pane |
| 3 | ViewNotesPage | The element is the Notes Page |
| 4 | ViewHandoutMaster | The element is the Handout Master |
| 5 | ViewNotesMaster | The element is the Notes Master |
| 6 | ViewOutline | The element is the Outline Pane |
| 7 | ViewSlideSorter | The element is the Slide Sorter |
| 8 | ViewTitleMaster | The element is the Title Master |
| 9 | ViewNormal | The element is the Normal View |
| 10 | ViewPrintPreview | The element is the Print Preview Pane |
| 11 | ViewThumbnails | The element is the Thumbnail Pane |
| 12 | ViewMasterThumbnails | The element is the Slide Master Thumbnail Pane |
| 13 | ViewNotesText | The element is the Notes Pane |
| 14 | ViewOutlineMaster | The element is the Slide Master Outline Pane |
| 15 | ViewSlideShow | The element is the Slide Show View |
| 16 | ViewSlideShowFullScreen | The element is the Full Screen Slide Show View |
| 17 | ViewSlideShowBrowse | The element is the Slide Show in a Window |
| 18 | ViewPresenterSlide | The element is the Presenter View Slide Pane |
| 19 | ViewPresenterNotes | The element is the Presenter View Notes Pane |
| 20 | ViewPresenterNextStep | The element is the Presenter View Next Step Pane |
| 21 | ViewPresenterTitle | The element is the Presenter View Title |
| 22 | ViewGridSections | The element is the Grid View Sections Pane |
| 23 | ViewGridThumbnails | The element is the Grid View Thumbnail Pane |
| 24 | ViewGridSectionTitle | The element is the Grid View Section Title |
| 25 | ViewGridThumbnailZoom | The element is the Grid View Thumbnail Pane Zoom Control |
| 26 | ViewGridBack | The element is the Grid View Back Button |
| 27 | ViewProtected | The element is the Protected View Window |
| 28 | ViewVisualBasic | The element is the Visual Basic Window |
| 29 | ViewNone | The element is not a content view |
| 30 | ViewUnknown | The element is not a known view type |

### AutomationId Patterns

PowerPoint elements typically use:
- Generic automation IDs from Office shared components
- Localized names (may vary by language)
- Ribbon element IDs follow Office common patterns

---

## 5. Comments Pane Accessibility

### Opening the Comments Pane

| Method | Shortcut |
|--------|----------|
| Ribbon Access Key | Alt+Z, C (Windows) |
| Direct Shortcut | Alt+R, P, P (to toggle) |
| Navigation | F6 to cycle to pane |

### Comments Pane UIA Structure

#### Linear Walkthrough

**Comments Pane (Pane) contains:**
1. **New Comment Button**
2. **Comment Threads Container** - Contains multiple Comment Threads
   - Each **Comment Thread** has:
     - **Parent Comment** - Contains Author Name, Timestamp, Comment Text, More Actions Button, Like Button
     - **Reply Comments** (collapsible) - Contains Reply 1, Reply 2, etc.
3. **Reply Input Field**

#### 2D Visual Map

```
Comments Pane (Pane)
├── New Comment Button
├── Comment Threads Container
│   ├── Comment Thread 1
│   │   ├── Parent Comment
│   │   │   ├── Author Name
│   │   │   ├── Timestamp
│   │   │   ├── Comment Text
│   │   │   ├── More Actions Button
│   │   │   └── Like Button
│   │   └── Reply Comments (collapsible)
│   │       ├── Reply 1
│   │       └── Reply 2
│   └── Comment Thread 2
│       └── ...
└── Reply Input Field
```

### Keyboard Navigation

| Action | Windows Shortcut | Mac Shortcut |
|--------|-----------------|--------------|
| Add new comment | Ctrl+Alt+M | Cmd+Shift+M |
| Post comment/reply | Ctrl+Enter | Cmd+Enter |
| Move between threads | Arrow Up/Down | Arrow Up/Down |
| Expand thread | Right Arrow | Right Arrow |
| Collapse thread | Left Arrow | Left Arrow |
| Move between elements | Tab / Shift+Tab | Tab / Shift+Tab |
| Toggle comment/anchor | Alt+F12 | Option+F12 |

### Stable UIAutomationId Identifiers (v0.0.30 Research)

The following UIAutomationId values were discovered through NVDA addon development testing (December 2025) and provide stable, non-localized identifiers for Comments pane elements:

| UIAutomationId | Element | Match Type | Notes |
|---------------|---------|------------|-------|
| `NewCommentButton` | New comment button | Exact | Always present in pane |
| `CommentsList` | Comments container | Exact | Role=20 (list) |
| `cardRoot_<N>_<GUID>` | Comment thread | Prefix | N=position, GUID=unique per comment |
| `firstPaneElement<GUID>` | Comments Pane root | Prefix | Role=56, name="Comments Pane" |

**Pattern Details:**
- `cardRoot_` format: `cardRoot_1_B0759BAC-813A-4A38-BF43-124049563ACD`
  - The `_1_` portion appears constant (may indicate nesting level)
  - GUID is unique per comment thread
- `firstPaneElement` format: `firstPaneElement5F6E8B9B-F2EF-4238-B749-89D6F5F132DD`
  - GUID appears stable within a session
  - Identifies the pane container in parent chain

**Usage in NVDA Addon:**
```python
# Check if focus is in Comments pane using stable UIAutomationId
uia_id = getattr(obj, 'UIAAutomationId', '') or ''
if (uia_id == 'NewCommentButton' or
    uia_id == 'CommentsList' or
    uia_id.startswith('cardRoot_') or
    uia_id.startswith('firstPaneElement')):
    # Focus is inside Comments pane
    return True
```

**Advantages over name-based detection:**
- Not dependent on localized text (works in any language)
- More reliable than checking for "comment" in element names
- Survives Office UI updates that may change display text

### Comment Card Announcement Reformatting (v0.0.37+)

**Problem:** The default comment card announcement is verbose:
```
"Comment thread started by Brett Humphrey, with 1 reply"
```

**Solution:** Reformat to concise author + comment text:
- Unresolved: `"Brett Humphrey: @John Smith please review..."`
- Resolved: `"Resolved - Brett Humphrey: Should we add a note..."`
- Reply: `"Reply - John Smith: Looks good to me"`

**Technical Implementation (v0.0.37 - Cancel and Re-announce):**

After testing, `event_NVDAObject_init` does NOT work for UIA objects in PowerPoint - properties
aren't available at init time. The working solution uses `event_gainFocus` with cancel-and-reannounce:

```python
import speech

def event_gainFocus(self, obj, nextHandler):
    """Cancel default announcement and speak reformatted version."""
    uia_id = getattr(obj, 'UIAAutomationId', '') or ''
    name = getattr(obj, 'name', '') or ''
    description = getattr(obj, 'description', '') or ''

    is_comment_card = (
        uia_id.startswith('cardRoot_') or
        'Comment thread started by' in name
    )

    if is_comment_card:
        is_resolved = name.startswith("Resolved ")

        # Extract author from "Comment thread started by Author"
        if " started by " in name:
            author_part = name.split(" started by ", 1)[1]
            if ", with " in author_part:
                author = author_part.split(", with ")[0]
            else:
                author = author_part

            if author and description:
                speech.cancelSpeech()  # Stop default announcement
                if is_resolved:
                    ui.message(f"Resolved - {author}: {description}")
                else:
                    ui.message(f"{author}: {description}")
                return  # Don't call nextHandler

    nextHandler()
```

**Why cancel-and-reannounce works:**
1. `event_gainFocus` fires when comment gets focus
2. `speech.cancelSpeech()` stops NVDA's queued verbose announcement
3. `ui.message()` queues our concise reformatted message
4. By returning early, we prevent default processing

**Why event_NVDAObject_init failed (v0.0.36):**
- `event_NVDAObject_init` is NOT called for UIA objects in PowerPoint
- Properties like `UIAAutomationId` may not be available at object init time
- Only works reliably for certain object types

### Avoiding Speech Cutoff During Slide Navigation (v0.0.48)

**Problem:** When using PageUp/PageDown in Comments pane to navigate slides, the slide title
announcement (from worker thread) was being cut off by `speech.cancelSpeech()` in comment
reformatting.

**Solution:** Track navigation state with `_just_navigated` flag:

```python
# In PageUp/PageDown handler - set flag before navigation
self._pending_auto_focus = True
self._worker.request_navigate(direction)

# When focus returns to comments pane after navigation
if is_in_comments and getattr(self, '_pending_auto_focus', False):
    self._pending_auto_focus = False
    self._just_navigated = True  # Skip cancelSpeech for first comment

# In comment reformatting - check flag before canceling speech
if author and description:
    if not getattr(self, '_just_navigated', False):
        speech.cancelSpeech()  # Normal case - cancel verbose announcement
    else:
        self._just_navigated = False  # Clear flag, don't cancel (let title finish)
    ui.message(formatted)
```

**Result:** User hears complete sequence: `"3: Market Analysis"` → `"Has 2 comments"` → `"Author: comment text"`

**Comment Card Properties (from v0.0.32 research):**

| Property | Content |
|----------|---------|
| `name` | "Comment thread started by Author" or "Resolved comment thread started by Author" |
| `description` | The actual comment text (e.g., "@John Smith please review the title") |
| `UIAAutomationId` | `cardRoot_1_<GUID>` |
| `role` | 21 (list item) |
| `states` | FOCUSABLE, COLLAPSED, FOCUSED |

### Screen Reader Behavior

- **NVDA:** Reads comments automatically when focused
- **JAWS:** Reads with Control Description option enabled
- **Narrator:** Use SR+0 to read first comment
- Threaded replies accessible via Up/Down within thread
- Author and timestamp announced with comment text

---

## 6. Known Limitations and Workarounds

### Major UIA Limitations in PowerPoint

#### 1. Slide Content Not Exposed via UIA
- **Issue:** Shapes, text boxes, and slide content not accessible through standard UIA
- **Workaround:** Use PowerPoint COM automation (IDispatch) to access object model
- **NVDA Solution:** The appModule adds `paneClassDC` to `badUIAWindowClasses`

#### 2. Slideshow Mode Object Detection (NVDA Issue #16161)
- **Issue:** NVDA inconsistently detects grouped objects, shapes, and SmartArt in slideshow mode
- **Affected Elements:**
  - Grouped text with decorative shapes (alt text ignored)
  - Grouped shapes with alt text (completely ignored)
  - Individual shapes (inconsistent detection)
  - SmartArt (alt text reads, internal text ignored)
- **JAWS Behavior:** Correctly recognizes all objects (expected behavior)
- **Status:** Requires NVDA appModule fixes

#### 3. PowerPoint 2013+ UIA Implementation Issues (NVDA Issue #3578)
- **Issue:** Microsoft's incomplete UIA implementation broke existing assistive technology workarounds
- **Symptom:** Tab navigation announces "text box" instead of content
- **Fix Applied:** Add `mdiClass` to `badUIAWindowClasses` in PowerPoint appModule

#### 4. Text Editing in Placeholders (NVDA Issue #11094)
- **Issue:** Cannot navigate text with arrow keys in placeholders
- **Cause:** UIA/COM interaction issues
- **Status:** Fixed in NVDA via TextInfo improvements

#### 5. Wide Characters and Caret Position (NVDA PR #17015)
- **Issue:** Caret position reporting fails with wide characters (emoji, CJK)
- **Fix:** Improved TextInfo implementation for character offset handling

#### 6. UIA Events Not Captured
- **Issue:** UI Automation events from PowerPoint may not be catchable programmatically
- **Example:** Ribbon tab selection events visible in AccEvent but not in code
- **Workaround:** Use COM event handlers instead

### General Accessibility Limitations

| Limitation | Impact | Workaround |
|-----------|--------|------------|
| Nested tables confuse screen readers | Lose cell count | Avoid nested tables |
| Reading order may not match visual | Content read out of order | Use Reading Order pane |
| SmartArt internal text not accessible | Content missed | Add alt text to whole object |
| Empty text boxes flagged | False positives | Recent update: no longer flagged |
| Mathematical equations | May not be fully readable | Use MathML where possible |

---

## 7. Exploration Tools and Techniques

### Recommended Tools

#### 1. Accessibility Insights for Windows (Recommended)
- Modern Microsoft tool replacing legacy inspect/accevent
- Download: https://accessibilityinsights.io/
- Features:
  - Live Inspect mode
  - Automated accessibility testing
  - UIA property and pattern display
  - Event monitoring
- **Known Limitation:** UIA tree navigation shows only ancestors, not siblings
- **Workaround:** Use Ctrl+Shift+F7/F8 for navigation, or fall back to Inspect.exe

#### 2. Inspect.exe (Windows SDK)
- Location: `C:\Program Files (x86)\Windows Kits\10\bin\<version>\<platform>\inspect.exe`
- Features:
  - Raw, Control, and Content tree views
  - UIA and MSAA modes
  - Full tree navigation
- **Tip:** Switch to Control View (not Raw View) for practical element access

#### 3. FlaUInspect
- Download: https://github.com/FlaUI/FlaUInspect/releases
- Features:
  - UIA2 and UIA3 mode selection
  - XPath display for elements
  - Control highlighting
- **Note:** May show different elements than Inspect.exe due to FlaUI library differences

#### 4. AccEvent (Windows SDK)
- Location: Windows SDK `\bin\<version>\<platform>\Accevent.exe`
- Features:
  - Real-time event monitoring
  - FocusChanged, SelectionItem, PropertyChanged events
  - Scope filtering
- **Tip:** Scope to specific elements to reduce noise

### Python Libraries for Exploration

#### pywinauto (Recommended)
```python
from pywinauto import Application
app = Application(backend="uia").connect(path="POWERPNT.EXE")
app.window().dump_tree()  # Print UIA tree
```

#### Python-UIAutomation-for-Windows
```python
import uiautomation as auto
# Find PowerPoint window
ppt = auto.WindowControl(searchDepth=1, ClassName='PPTFrameClass')
# Enumerate children
for control in ppt.GetChildren():
    print(control.ControlTypeName, control.Name)
```

#### comtypes with UIAutomationCore
```python
from comtypes.client import GetModule, CreateObject
GetModule('UIAutomationCore.dll')
from comtypes.gen.UIAutomationClient import CUIAutomation, IUIAutomation
uia = CreateObject(CUIAutomation._reg_clsid_, interface=IUIAutomation)
root = uia.GetRootElement()
```

---

## 8. Code Examples

### Accessing PowerPoint via COM (Python)

```python
import win32com.client

# Connect to running PowerPoint instance
ppt = win32com.client.Dispatch("PowerPoint.Application")

# Access active presentation
presentation = ppt.ActivePresentation

# Get current slide
slide = ppt.ActiveWindow.View.Slide

# Enumerate shapes on slide
for shape in slide.Shapes:
    print(f"Shape: {shape.Name}, Type: {shape.Type}")
    if shape.HasTextFrame:
        if shape.TextFrame.HasText:
            print(f"  Text: {shape.TextFrame.TextRange.Text}")

# Access comments
for comment in slide.Comments:
    print(f"Comment by {comment.Author}: {comment.Text}")
```

### Reading ViewType Custom Property (C#)

```csharp
using System.Windows.Automation;

// ViewType GUID
Guid viewTypeGuid = new Guid("{F065BAA7-2794-48B6-A927-193DA1540B84}");

// Register the property
AutomationProperty viewTypeProp = AutomationProperty.Register(
    viewTypeGuid, "ViewType", typeof(int));

// Get value from pane element
int viewType = (int)paneElement.GetCurrentPropertyValue(viewTypeProp);
```

### Finding PowerPoint Elements with pywinauto

```python
from pywinauto import Application

app = Application(backend="uia").connect(path="POWERPNT.EXE")
main_window = app.window(class_name="PPTFrameClass")

# Access ribbon
ribbon = main_window.child_window(control_type="Pane", title_re="Ribbon.*")

# Access thumbnail pane
thumbnails = main_window.child_window(auto_id="ThumbnailPane")

# Access comments pane (if open)
comments = main_window.child_window(title="Comments", control_type="Pane")
```

### NVDA-Style COM Event Handling (Python)

```python
import win32com.client
import pythoncom

class PowerPointEvents:
    def OnWindowSelectionChange(self, sel):
        """Fires when selection changes in PowerPoint window"""
        print(f"Selection changed: {sel}")

    def OnSlideShowNextSlide(self, wn):
        """Fires when slideshow advances to next slide"""
        print(f"Slide show advanced: {wn.View.Slide.SlideIndex}")

# Connect with events
ppt = win32com.client.DispatchWithEvents(
    "PowerPoint.Application",
    PowerPointEvents
)

# Keep message loop running
pythoncom.PumpMessages()
```

### Getting Slide Title via COM (Python)

The slide title can be accessed through the `Shapes.Title` property. This is useful for
announcing slide context during navigation.

```python
import win32com.client

ppt = win32com.client.Dispatch("PowerPoint.Application")
window = ppt.ActiveWindow
slide = window.View.Slide

# Check if slide has a title placeholder
if slide.Shapes.HasTitle:
    title_shape = slide.Shapes.Title
    if title_shape.HasTextFrame:
        text_frame = title_shape.TextFrame
        if text_frame.HasText:
            title_text = text_frame.TextRange.Text.strip()
            print(f"Slide title: {title_text}")
else:
    print("Slide has no title placeholder")

# Get slide number and total
slide_index = slide.SlideIndex
total_slides = window.Presentation.Slides.Count
print(f"Slide {slide_index} of {total_slides}")
```

**Key COM properties for slide title access:**
- `Shapes.HasTitle` - Boolean, True if slide has a title placeholder
- `Shapes.Title` - Returns the Shape object representing the title
- `Shape.HasTextFrame` - Boolean, True if shape contains text
- `TextFrame.HasText` - Boolean, True if text frame has content
- `TextFrame.TextRange.Text` - The actual title text

**References:**
- [Shapes.Title property (PowerPoint)](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.Title)
- [Shapes.HasTitle property (PowerPoint)](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.HasTitle)

### Getting Slide Notes via COM (Python)

Slide notes are accessed through the `NotesPage` property. The notes page contains multiple
shapes - Placeholder(1) is the slide thumbnail, Placeholder(2) is the notes body text.

```python
import win32com.client

ppt = win32com.client.Dispatch("PowerPoint.Application")
window = ppt.ActiveWindow
slide = window.View.Slide

# Access notes via NotesPage
notes_page = slide.NotesPage

# Placeholder 2 is the notes body text
placeholder = notes_page.Shapes.Placeholders(2)
if placeholder.HasTextFrame:
    text_frame = placeholder.TextFrame
    if text_frame.HasText:
        notes_text = text_frame.TextRange.Text.strip()
        print(f"Notes: {notes_text}")
else:
    print("No notes on this slide")
```

**Key COM properties for slide notes access:**
- `Slide.NotesPage` - Returns SlideRange representing the notes page
- `NotesPage.Shapes.Placeholders(2)` - The notes body placeholder (index 2)
- `Shape.HasTextFrame` - Boolean, True if shape contains text frame
- `TextFrame.HasText` - Boolean, True if text frame has content
- `TextFrame.TextRange.Text` - The actual notes text

**NotesPage placeholder structure:**
| Placeholder Index | Content |
|-------------------|---------|
| 1 | Slide thumbnail image |
| 2 | Notes body text |

**References:**
- [Slide.NotesPage property (PowerPoint)](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.notespage)
- [TextFrame object (PowerPoint)](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.textframe)
- [Shape.TextFrame property (PowerPoint)](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.textframe)

---

## 9. NVDA PowerPoint AppModule Analysis

### Source Location
`nvda/source/appModules/powerpnt.py`

### Key Design Decisions

#### 1. Disabled UIA Window Classes
NVDA disables UIA for certain PowerPoint window classes:
- `paneClassDC` - Main slide editing area
- `mdiClass` - MDI container
- `screenClass` - Slideshow mode

These are added to `badUIAWindowClasses` to fall back to COM automation.

#### 2. COM Event Handling
The appModule defines `EApplication` COM interface for:
- `WindowSelectionChange` - Selection tracking
- `SlideShowNextSlide` - Slideshow navigation

#### 3. TextInfo Implementation
`TextFrameTextInfo` extends `OffsetsTextInfo` to work with PowerPoint's `TextRange` COM object:
- Character, word, line, paragraph offset retrieval
- Caret position and selection management
- Format field extraction (fonts, styles, colors, links, bullets)
- Bounding rectangle calculations

#### 4. Shape Handling
Shapes are mapped from `msoShapeTypes` to NVDA roles. The module:
- Detects overlapping shapes
- Calculates edge distances
- Reports off-slide positioning
- Exposes action button actions via name property

#### 5. Slideshow Mode
`ReviewableSlideshowTreeInterceptor` provides:
- Document-style navigation without traditional caret
- `SlideShowTreeInterceptorTextInfo` for continuous text exposure
- MathML field extraction support

### Recent Improvements (2024-2025)
- PR #17015: Fixed wide character caret reporting
- PR #17004: Added braille display routing key support
- Commit 0101cc4: Enhanced shape type reporting with localizable labels

---

## 10. References

### Microsoft Documentation

- [PowerPoint Custom Properties](https://learn.microsoft.com/en-us/office/uia/powerpoint/powerpointcustomproperties)
- [PowerPoint Enumerations](https://learn.microsoft.com/en-us/office/uia/powerpoint/powerpointenumerations)
- [UI Automation Custom Extensions in Office](https://learn.microsoft.com/en-us/office/uia/)
- [UI Automation Tree Overview](https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-treeoverview)
- [Inspect.exe Documentation](https://learn.microsoft.com/en-us/windows/win32/winauto/inspect-objects)
- [AccEvent Documentation](https://learn.microsoft.com/en-us/windows/win32/winauto/accessible-event-watcher)
- [Screen Reader Support for PowerPoint](https://support.microsoft.com/en-us/office/screen-reader-support-for-powerpoint-9d2b646d-0b79-4135-a570-b8c7ad33ac2f)
- [Use Keyboard Shortcuts to Navigate Modern Comments](https://support.microsoft.com/en-us/topic/use-keyboard-shortcuts-to-navigate-modern-comments-in-powerpoint-e6924fd8-43f2-474f-a1c5-7ccdfbf59b3b)

### NVDA Resources

- [NVDA PowerPoint AppModule Source](https://github.com/nvaccess/nvda/blob/master/source/appModules/powerpnt.py)
- [NVDA Issue #3578](https://github.com/nvaccess/nvda/issues/3578) - PowerPoint 2013 UIA issues
- [NVDA Issue #4850](https://github.com/nvaccess/nvda/issues/4850) - Slideshow reading
- [NVDA Issue #16161](https://github.com/nvaccess/nvda/issues/16161) - Object detection in slideshow
- [NVDA PR #17015](https://github.com/nvaccess/nvda/pull/17015) - Wide character fix
- [NVDA Developer Guide](https://www.nvaccess.org/files/nvda/documentation/developerGuide.html)

### Tools

- [Accessibility Insights for Windows](https://accessibilityinsights.io/docs/windows/overview/)
- [FlaUInspect](https://github.com/FlaUI/FlaUInspect)
- [Python-UIAutomation-for-Windows](https://github.com/yinkaisheng/Python-UIAutomation-for-Windows)
- [pywinauto Documentation](https://pywinauto.readthedocs.io/en/latest/)

### Community Discussions

- [Stack Overflow: PowerPoint UIA Events](https://stackoverflow.com/questions/14881705/ui-automation-events-not-getting-caught-from-powerpoint-2007)
- [Stack Overflow: PodiumParent Panel Access](https://stackoverflow.com/questions/64117225/powerpnt-exe-ui-automation-get-children-of-podiumparent-panel-using-inspect-exe)
- [Microsoft 365 Developer Blog: UIA Custom Properties](https://devblogs.microsoft.com/microsoft365dev/use-ui-automation-custom-properties-to-customize-your-assistive-technologies-to-office-applications/)

---

## Document Information

- **Created:** December 4, 2025
- **Purpose:** NVDA Plug-in Development for PowerPoint Accessibility
- **Status:** Complete initial research
- **Next Steps:** Use findings to design plug-in architecture and identify enhancement opportunities

---

## Appendix A: UIA Tree Exploration Script

```python
"""
PowerPoint UIA Tree Explorer
Requires: pywinauto, comtypes
"""
import sys
from pywinauto import Application
from pywinauto.findwindows import ElementNotFoundError

def explore_powerpoint():
    try:
        app = Application(backend="uia").connect(path="POWERPNT.EXE")
    except ElementNotFoundError:
        print("PowerPoint is not running")
        sys.exit(1)

    main_window = app.window(class_name="PPTFrameClass")

    print("=== PowerPoint UIA Tree ===\n")

    def print_tree(element, depth=0):
        indent = "  " * depth
        try:
            ctrl_type = element.element_info.control_type or "Unknown"
            name = element.element_info.name or "(no name)"
            auto_id = element.element_info.automation_id or ""
            class_name = element.element_info.class_name or ""

            print(f"{indent}[{ctrl_type}] {name}")
            if auto_id:
                print(f"{indent}  AutomationId: {auto_id}")
            if class_name:
                print(f"{indent}  ClassName: {class_name}")

            for child in element.children():
                print_tree(child, depth + 1)
        except Exception as e:
            print(f"{indent}Error: {e}")

    print_tree(main_window, 0)

if __name__ == "__main__":
    explore_powerpoint()
```

## Appendix B: ViewType Property Reader

```python
"""
Read PowerPoint ViewType Custom Property
Requires: comtypes
"""
import comtypes
from comtypes import GUID
from comtypes.client import GetModule, CreateObject

# Load UIAutomation type library
GetModule('UIAutomationCore.dll')
from comtypes.gen.UIAutomationClient import (
    CUIAutomation, IUIAutomation, IUIAutomationElement
)

# ViewType Property GUID
VIEWTYPE_GUID = GUID("{F065BAA7-2794-48B6-A927-193DA1540B84}")

# View type names
VIEW_TYPES = {
    1: "ViewSlide", 2: "ViewSlideMaster", 3: "ViewNotesPage",
    4: "ViewHandoutMaster", 5: "ViewNotesMaster", 6: "ViewOutline",
    7: "ViewSlideSorter", 8: "ViewTitleMaster", 9: "ViewNormal",
    10: "ViewPrintPreview", 11: "ViewThumbnails", 12: "ViewMasterThumbnails",
    13: "ViewNotesText", 14: "ViewOutlineMaster", 15: "ViewSlideShow",
    16: "ViewSlideShowFullScreen", 17: "ViewSlideShowBrowse",
    18: "ViewPresenterSlide", 19: "ViewPresenterNotes",
    20: "ViewPresenterNextStep", 21: "ViewPresenterTitle",
    22: "ViewGridSections", 23: "ViewGridThumbnails",
    24: "ViewGridSectionTitle", 25: "ViewGridThumbnailZoom",
    26: "ViewGridBack", 27: "ViewProtected", 28: "ViewVisualBasic",
    29: "ViewNone", 30: "ViewUnknown"
}

def get_viewtype(element):
    """Get ViewType custom property from a PowerPoint pane element"""
    try:
        # Note: Actual implementation requires registering custom property
        # This is a simplified conceptual example
        value = element.GetCurrentPropertyValue(VIEWTYPE_GUID)
        return VIEW_TYPES.get(value, f"Unknown ({value})")
    except Exception as e:
        return f"Error: {e}"

if __name__ == "__main__":
    uia = CreateObject(CUIAutomation._reg_clsid_, interface=IUIAutomation)
    # Further implementation would find PowerPoint panes and read ViewType
    print("ViewType property reader initialized")
```
