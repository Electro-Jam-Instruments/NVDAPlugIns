# NVDA Native PowerPoint Support - Deep Research Analysis

## Executive Summary

This document provides a comprehensive analysis of how NVDA (NonVisual Desktop Access) handles Microsoft PowerPoint natively. The research covers NVDA's source code architecture, COM vs UIA decision logic, known issues, and extension points for our plugin development.

### Key Findings

1. NVDA uses COM automation (PowerPoint Object Model) as primary accessibility mechanism, deliberately disabling incomplete UIA support
2. The powerpnt.py app module is approximately 1500+ lines of specialized code
3. Core limitation: NVDA has no native support for PowerPoint comments/annotations
4. Extension is possible via app module overlay classes and global plugins
5. Recent improvements (2024) focused on text navigation accuracy, not comment support

---

## 1. NVDA PowerPoint App Module Architecture

### 1.1 File Location and Structure

**Primary File:** `source/appModules/powerpnt.py`

The module defines a comprehensive class hierarchy for PowerPoint accessibility:

```
AppModule (powerpnt.py)
    |
    +-- PaneClassDC (Base window handler)
    |       |
    |       +-- DocumentWindow (Presentation document)
    |       |
    |       +-- SlideShowWindow (Slideshow mode)
    |
    +-- PpObject (Foundation class)
    |       |
    |       +-- SlideBase
    |       |       +-- Slide
    |       |       +-- Master
    |       |
    |       +-- Shape
    |       |       +-- ChartShape
    |       |       +-- Table
    |       |
    |       +-- TextFrame
    |       |       +-- TableCellTextFrame
    |       |       +-- NotesTextFrame
    |       |
    |       +-- TableCell
    |
    +-- SlideShowTreeInterceptor
    |       +-- ReviewableSlideshowTreeInterceptor
    |
    +-- TextFrameTextInfo (Text navigation)
```

### 1.2 COM Interface Integration

NVDA establishes COM event handling through the `EApplication` interface:

```python
# COM Event Sink Class
class ppEApplicationSink:
    # Handles WindowSelectionChange events
    # Handles SlideShowNextSlide events
```

**Key COM Events Monitored:**
- `WindowSelectionChange`: Triggers focus update when selection changes in PowerPoint
- `SlideShowNextSlide`: Handles slide advancement during presentations

### 1.3 Key Classes Explained

#### PaneClassDC
- Base window handler fetching PowerPoint object model access
- Provides access to current slides and application version
- Uses window class name "paneClassDC" for identification

#### DocumentWindow
- Represents the main presentation document
- Manages selection changes between slides, shapes, and text frames
- Bounces focus to appropriate child objects

#### PpObject
- Foundation class for slides, shapes, and text frames
- Handles keyboard navigation and selection management
- Base class for PowerPoint-specific accessibility objects

#### TextFrame and TextFrameTextInfo
- Provides editable text support
- Implements offset-based text navigation using PowerPoint's TextRange API
- Supports character, word, line, paragraph, and sentence navigation
- Extracts formatting information (fonts, colors, bold, italic, hyperlinks)

#### SlideShowTreeInterceptor
- Manages slideshow/presentation mode accessibility
- Inherits from DocumentTreeInterceptor for sayAll functionality
- Enables cursor-based navigation through slide content

---

## 2. COM vs UIA Decision Logic

### 2.1 The Problem with UIA in PowerPoint

Microsoft attempted to provide UI Automation support for PowerPoint, but as documented in NVDA Issue #3578:

> "Microsoft has now tried to provide an accessibility implementation for PowerPoint using UI Automation. But as usual its far from complete, yet at the same time cripples any existing support/hacks by other ATs."

### 2.2 NVDA's Solution: Disable UIA for PowerPoint

NVDA explicitly disables UIA for specific PowerPoint window classes by adding them to `badUIAWindowClasses`:

**Disabled Window Classes:**
- `paneClassDC` - Main content pane
- `mdiClass` - MDI container window (PowerPoint 2013+)
- `screenClass` - Slideshow presentation screen

### 2.3 Implementation Details

The AppModule's `chooseNVDAObjectOverlayClasses` method checks window class names:

```python
def chooseNVDAObjectOverlayClasses(self, obj, clsList):
    windowClass = obj.windowClassName
    if windowClass in ("paneClassDC", "mdiClass"):
        # Apply PowerPoint-specific overlay classes
        # Use COM automation instead of UIA
```

**Critical Note:** Adding `mdiClass` to global `badUIAWindowClasses` was rejected because:
- `mdiClass` is a generic MDI container window used by multiple Office apps
- Solution must be limited to PowerPoint appModule scope

### 2.4 Fallback Behavior

When UIA is disabled, NVDA falls back to:
1. PowerPoint COM Object Model (primary)
2. IAccessible/MSAA (secondary)
3. Window messages (tertiary)

---

## 3. What NVDA Handles Natively

### 3.1 Feature Matrix

| Feature | Support Level | Implementation Method |
|---------|--------------|----------------------|
| Focus tracking (slides/shapes) | Full | COM + Events |
| Text box reading | Full | TextFrameTextInfo |
| Text editing/caret | Full (improved 2024) | COM TextRange API |
| Shape navigation | Full | COM Selection API |
| Shape announcements | Full | Role mapping |
| Table navigation | Full | TableCell class |
| Table structure (rows/cols) | Full | COM Table API |
| Slide show mode | Full | SlideShowTreeInterceptor |
| Speaker notes | Full | NotesTextFrame |
| Notes pane | Full | Custom gesture |
| Outline view | Partial | Known issues |
| Hyperlinks | Full | TextRange links |
| MathType equations | Full | Requires MathType |
| Charts | Basic | Entry interaction |
| SmartArt | Limited | Known issues |
| Grouped objects | Limited | Alt text issues |
| Comments/Annotations | NONE | Not implemented |
| Protected view | NONE | Issue #3007 |

### 3.2 Keyboard Gestures Defined

NVDA's PowerPoint module defines these gesture bindings:

**Navigation:**
- `Tab` / `Shift+Tab`: Selection navigation between shapes
- Arrow Keys: Shape movement with location announcements
- `Page Up` / `Page Down`: Slide navigation
- `Home` / `End`: First/last slide

**Editing:**
- `Enter` / `F2`: Enter shape/text editing mode
- `Escape`: Return to document/exit editing

**Slideshow Mode:**
- `Space` / `Backspace`: Advance/reverse slides
- `Ctrl+Shift+S`: Toggle speaker notes reading
- Arrow Keys: Navigate slide content

**Shape Position:**
- `Shift+Arrow`: Report shape location relative to slide edges

### 3.3 Event Handling

Events NVDA monitors in PowerPoint:

| Event | Handler | Purpose |
|-------|---------|---------|
| gainFocus | event_gainFocus | Track focus changes |
| stateChange | event_stateChange | Monitor state changes |
| valueChange | event_valueChange | Track value updates |
| WindowSelectionChange | COM Event | Selection changes |
| SlideShowNextSlide | COM Event | Slide advancement |

---

## 4. Known Issues and Limitations

### 4.1 GitHub Issues Summary

| Issue # | Title | Status | Impact |
|---------|-------|--------|--------|
| #7288 | Links in slideshow not announced on Tab | Open | High |
| #16161 | Objects/groups not perceived in slideshow | Closed | Medium |
| #12719 | Outline view reads "linefeed" | Open | Medium |
| #3578 | PowerPoint 2013 announces "text box" | Fixed 2014.1 | Historical |
| #3007 | Protected view not accessible | Open | High |
| #15677 | Error sounds with visual highlighter | Open | Low |
| #17167 | Grant access dialog not announced | Open | Medium |
| #4850 | Slideshow auto-reading | Open | Medium |

### 4.2 Detailed Issue Analysis

#### Links in Slideshow Mode (#7288)
- Tab navigates between links but NVDA reads nothing
- Focus changes aren't being captured properly
- Compare: JAWS announces links correctly

#### Objects and Groups (#16161)
- Grouped shapes with alt text are ignored
- Individual shapes inconsistently recognized
- SmartArt text inside objects not read
- Tab/Shift+Tab produces no results for some objects

#### Outline View (#12719)
- NVDA reads only "linefeed" and "blank"
- Narrator reads correctly (uses UIA?)
- COM automation may not expose outline text properly

### 4.3 Recent Improvements (2024)

**PR #17015 - TextInfo Implementation Overhaul:**

Fixed Issues:
- Wide character caret position reporting (#17006)
- Inaccurate text position with NVDA+Delete (#9941)

Changes:
- Added `_getCharacterOffsets()` for precise character navigation
- Added `_getWordOffsets()` for word boundary detection
- Added `_getLineNumFromOffset()` for line calculations
- Added `_getParagraphOffsets()` for paragraph boundaries
- Improved `_getPptTextRange()` validation and error handling

---

## 5. Extension Points for Our Plugin

### 5.1 Available Hook Points

#### App Module Methods
```python
# Override in custom appModule
def chooseNVDAObjectOverlayClasses(self, obj, clsList):
    """Add custom NVDAObject classes"""

def event_gainFocus(self, obj, nextHandler):
    """Intercept focus events"""

def event_NVDAObject_init(self, obj):
    """Modify object properties during initialization"""
```

#### Extension Points
```python
# From extensionPoints module
post_configProfileSwitch  # Config changes
filter_speechSequence     # Modify speech output
decide_executeGesture     # Intercept gestures
treeInterceptorHandler.post_browseModeStateChange  # Mode changes
```

### 5.2 Custom NVDAObject Overlay Classes

To add comment support, we can create overlay classes:

```python
class CommentShape(Shape):
    """Custom class for comment indicators"""

    def _get_name(self):
        # Extract comment text from COM
        pass

    role = controlTypes.Role.COMMENT  # Custom role
```

Register via `chooseNVDAObjectOverlayClasses`:
```python
def chooseNVDAObjectOverlayClasses(self, obj, clsList):
    if self._isCommentShape(obj):
        clsList.insert(0, CommentShape)
```

### 5.3 Adding Custom Gestures/Scripts

**CRITICAL UPDATE (v0.0.9):** Use the EXACT NVDA documentation pattern.
See `decisions.md` Decision 6 for the verified pattern.

```python
# CORRECT: NVDA documentation pattern (verified working v0.0.9)
from nvdaBuiltin.appModules.powerpnt import *

class AppModule(AppModule):  # Inherits from just-imported AppModule

    @script(
        description=_("Read comments on current slide"),
        category="PowerPoint",
        gestures=["kb:NVDA+shift+c"]
    )
    def script_readComments(self, gesture):
        # Implementation
        pass

    @script(
        description=_("Navigate to next comment"),
        gestures=["kb:NVDA+alt+c"]
    )
    def script_nextComment(self, gesture):
        # Implementation
        pass
```

### 5.4 Accessing PowerPoint COM from Plugin

```python
import comtypes.client

def getCommentsFromSlide(slide):
    """Extract comments from a PowerPoint slide via COM"""
    comments = []
    try:
        # Access Comments collection
        commentCollection = slide.Comments
        for i in range(1, commentCollection.Count + 1):
            comment = commentCollection.Item(i)
            comments.append({
                'author': comment.Author,
                'text': comment.Text,
                'datetime': comment.DateTime,
                'authorInitials': comment.AuthorInitials
            })
    except Exception:
        pass
    return comments
```

---

## 6. Recommendations for Our Implementation

### 6.1 Option 1: App Module Extension (Recommended)

**Approach:** Create a custom app module that extends NVDA's built-in PowerPoint support

**Pros:**
- Integrates seamlessly with existing NVDA PowerPoint code
- Can use existing COM infrastructure
- Inherits all existing functionality
- Users get enhanced experience without conflicts

**Cons:**
- Must name file `powerpnt.py` (shadows built-in)
- Need to re-implement or import base functionality
- Updates to NVDA core require plugin updates

**Implementation Path:**
1. Create `appModules/powerpnt.py` in addon
2. Import from NVDA's powerpnt module
3. Add comment-related overlay classes
4. Add custom scripts and gestures

### 6.2 Option 2: Global Plugin with Extension Points

**Approach:** Create a global plugin that hooks into PowerPoint via extension points

**Pros:**
- Doesn't shadow built-in app module
- Cleaner separation of concerns
- Easier maintenance across NVDA versions
- Can be enabled/disabled independently

**Cons:**
- More complex to integrate with existing PowerPoint handling
- May miss some PowerPoint-specific events
- Need to carefully manage focus and event flow

**Implementation Path:**
1. Create `globalPlugins/pptComments.py`
2. Use `chooseNVDAObjectOverlayClasses` for overlay classes
3. Register app module extension via `registerExecutableWithAppModule`
4. Implement comment detection and announcement

### 6.3 Option 3: Hybrid Approach (Most Flexible)

**Approach:** Global plugin for coordination + App module overlay for PowerPoint-specific handling

**Pros:**
- Best of both approaches
- Maximum flexibility
- Clean architecture
- Easy to maintain

**Cons:**
- More complex initial setup
- Two components to manage

**Recommendation:** Start with Option 1 for fastest results, refactor to Option 3 if maintenance becomes difficult.

---

## 7. Technical Reference

### 7.1 PowerPoint COM Object Model for Comments

```
Application
    +-- ActivePresentation
            +-- Slides (collection)
                    +-- Slide
                            +-- Comments (collection)
                                    +-- Comment
                                            - Author
                                            - AuthorIndex
                                            - AuthorInitials
                                            - DateTime
                                            - Text
                                            - Left, Top (position)
                                            +-- Replies (collection)
                                                    +-- Comment (reply)
```

### 7.2 Window Class Names

| Class Name | Description | UIA Status |
|------------|-------------|------------|
| paneClassDC | Main content pane | Disabled |
| mdiClass | MDI container | Disabled |
| screenClass | Slideshow screen | Disabled |
| NetUIHWND | Ribbon/UI elements | Enabled |

### 7.3 Shape Types Relevant to Comments

PowerPoint doesn't represent comments as shapes in the traditional sense. Comments are accessed via:
- `Slide.Comments` collection (COM)
- `Shape` objects with `msoShapeComment` type (older versions)
- Modern comments panel (UIA - partially accessible)

---

## 8. Sources and References

### Primary Sources
- [NVDA GitHub Repository](https://github.com/nvaccess/nvda)
- [NVDA Developer Guide](https://www.nvaccess.org/files/nvda/documentation/developerGuide.html)
- [NVDA User Guide - PowerPoint Section](https://www.nvaccess.org/files/nvda/documentation/userGuide.html)

### Key GitHub Issues
- [Issue #7288 - Links in Slideshow](https://github.com/nvaccess/nvda/issues/7288)
- [Issue #16161 - Objects/Groups in Slideshow](https://github.com/nvaccess/nvda/issues/16161)
- [Issue #3578 - PowerPoint 2013 Support](https://github.com/nvaccess/nvda/issues/3578)
- [Issue #3007 - Protected View](https://github.com/nvaccess/nvda/issues/3007)
- [Issue #12719 - Outline View](https://github.com/nvaccess/nvda/issues/12719)
- [PR #17015 - TextInfo Fix](https://github.com/nvaccess/nvda/pull/17015)

### Microsoft Documentation
- [PowerPoint VBA API Reference](https://learn.microsoft.com/en-us/office/vba/api/overview/powerpoint)
- [TextRange Object](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.textrange)
- [Application.WindowSelectionChange Event](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.Application.WindowSelectionChange)

### NVDA Add-on Development
- [NVDA Add-on Development Guide](https://github.com/nvdaaddons/DevGuide/wiki/NVDA-Add-on-Development-Guide)
- [NVDA Add-ons Directory](https://nvda-addons.org/)

---

## Document Information

- **Created:** December 4, 2025
- **Author:** Strategic Planning and Research Agent
- **Version:** 1.0
- **Purpose:** Foundation research for A11Y PowerPoint NVDA Plugin development
