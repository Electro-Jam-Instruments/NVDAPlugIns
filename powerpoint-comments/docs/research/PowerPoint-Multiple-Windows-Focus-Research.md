# PowerPoint COM: Reliably Getting the Focused Window with Multiple Open Presentations

## Executive Summary

This research addresses the critical issue of identifying the correct PowerPoint presentation window when multiple files are open simultaneously. The current implementation uses `Application.ActiveWindow`, which can return the wrong window when switching between presentations.

### Key Findings

1. **The Problem:** `ActiveWindow` returns the last-used window in the application, which may not be the window that triggered the `WindowSelectionChange` event
2. **The Solution:** Use the `sel` parameter's `Parent` property to get the correct window that triggered the event
3. **Alternative Approach:** Iterate `Application.Windows` collection and match by HWND to find the OS-focused window
4. **Best Practice:** Combine both approaches - use `sel.Parent` when available, fall back to HWND matching

### Recommended Implementation

Use the `WindowSelectionChange` event's `sel` parameter to get the correct window:

```python
def WindowSelectionChange(self, sel):
    """Event handler that receives the Selection object."""
    try:
        # Get the DocumentWindow that contains this selection
        # sel.Parent returns the DocumentWindow object
        window = sel.Parent

        # Now get the slide from THIS specific window
        slide = window.View.Slide
        slide_index = slide.SlideIndex

        # Or get the presentation
        presentation = window.Presentation
    except:
        # Fallback to ActiveWindow if sel.Parent fails
        slide = self._ppt_app.ActiveWindow.View.Slide
```

---

## 1. Understanding the PowerPoint COM Object Hierarchy

### 1.1 Object Model Structure

```
Application
    |
    +-- Windows (DocumentWindows collection)
    |       |
    |       +-- DocumentWindow(1)  (always the active window)
    |       +-- DocumentWindow(2)
    |       +-- DocumentWindow(n)
    |
    +-- ActiveWindow (DocumentWindow) - points to one of the above
    |
    +-- Presentations (collection)
            |
            +-- Presentation(1)
            |       |
            |       +-- Windows (DocumentWindows for this presentation)
            |
            +-- Presentation(2)
```

### 1.2 Key Object Relationships

| Object | Parent | Contains | Notes |
|--------|--------|----------|-------|
| Application | - | Windows, Presentations | Top-level |
| DocumentWindow | Application | Selection, View | Can view one presentation |
| Selection | DocumentWindow | ShapeRange, SlideRange, TextRange | What's selected |
| View | DocumentWindow | Slide, Type | Current view state |
| Presentation | Application | Slides, Windows | One .pptx file |

### 1.3 Important Properties

**DocumentWindow Properties:**
- `Parent` → Application
- `Presentation` → The Presentation being viewed
- `Selection` → Current Selection
- `View` → Current View
- `Active` → Boolean, True if window has focus
- `HWND` → Window handle (int)

**Selection Properties:**
- `Parent` → **DocumentWindow** (THIS IS THE KEY!)
- `Type` → ppSelectionType enum
- `SlideRange` → Selected slides
- `ShapeRange` → Selected shapes
- `TextRange` → Selected text

---

## 2. The Multi-Window Problem

### 2.1 Scenario

User has two PowerPoint files open:
- `Presentation1.pptx` - Window 1
- `Presentation2.pptx` - Window 2

User switches focus to Window 2 and navigates to slide 5.

### 2.2 Current Broken Code

```python
def WindowSelectionChange(self, sel):
    # PROBLEM: ActiveWindow might still point to Window 1
    current_slide = self._ppt_app.ActiveWindow.View.Slide
    # Returns slide from WRONG presentation!
```

**Why it fails:**
- `Application.ActiveWindow` is updated by PowerPoint but not always immediately
- During rapid window switching, `ActiveWindow` can lag behind
- COM event delivery timing doesn't guarantee `ActiveWindow` is updated first

### 2.3 Real-World Impact

1. User switches to `Presentation2.pptx`
2. `WindowSelectionChange` event fires
3. Plugin checks `ActiveWindow.View.Slide` → still pointing to Presentation1
4. Plugin announces comments from the WRONG presentation
5. User is confused and trust in the addon is broken

---

## 3. Solution 1: Use sel.Parent (RECOMMENDED)

### 3.1 The Solution

The `WindowSelectionChange` event passes a `Selection` object (`sel` parameter). This Selection object knows which window it belongs to.

**Object Model:**
```
WindowSelectionChange(sel)
    sel.Parent → DocumentWindow (the window that contains this selection)
        .View → View
            .Slide → The actual slide in THIS window
        .Presentation → The presentation in THIS window
```

### 3.2 Implementation

```python
class PowerPointEventSink(COMObject):
    _com_interfaces_ = [EApplication, IDispatch]

    def __init__(self, worker):
        super().__init__()
        self._worker = worker
        self._last_slide_by_window = {}  # Track per-window

    def WindowSelectionChange(self, sel):
        """Called when selection changes in ANY PowerPoint window.

        Args:
            sel: Selection object (IDispatch) - contains Parent property
        """
        try:
            log.debug("WindowSelectionChange event received")

            # Get the DocumentWindow that triggered this event
            # This is THE KEY - sel.Parent is the specific window
            window = sel.Parent

            if not window:
                log.warning("sel.Parent returned None - falling back to ActiveWindow")
                window = self._worker._ppt_app.ActiveWindow

            # Get slide from THIS specific window
            try:
                slide = window.View.Slide
                slide_index = slide.SlideIndex
                presentation = window.Presentation

                # Create unique key for this window
                # Use presentation name + window index
                window_key = f"{presentation.Name}_{window.WindowNumber}"

                # Check if slide changed in THIS window
                last_slide = self._last_slide_by_window.get(window_key, -1)

                if slide_index != last_slide:
                    log.info(f"Slide changed in {presentation.Name} to {slide_index}")
                    self._last_slide_by_window[window_key] = slide_index

                    # Pass both window and slide to worker
                    self._worker.on_slide_changed_event(
                        window=window,
                        slide_index=slide_index,
                        presentation_name=presentation.Name
                    )

            except Exception as e:
                log.debug(f"Could not get slide from window - {e}")

        except Exception as e:
            log.error(f"Error in WindowSelectionChange - {e}")
```

### 3.3 Updated Worker Method

```python
class PowerPointWorker:
    def on_slide_changed_event(self, window, slide_index, presentation_name):
        """Called by event sink when slide changes.

        Args:
            window: DocumentWindow that triggered the event
            slide_index: New slide index (1-based)
            presentation_name: Name of the presentation
        """
        log.info(f"Slide change in '{presentation_name}' - slide {slide_index}")

        # Use the SPECIFIC window passed from the event
        self._announce_slide_comments(window=window)

    def _announce_slide_comments(self, window=None):
        """Announce comment status for current slide.

        Args:
            window: Specific DocumentWindow to use (or None for ActiveWindow)
        """
        if window is None:
            window = self._ppt_app.ActiveWindow

        comments = self._get_comments_on_slide(window)

        if not comments:
            self._announce("No comments")
        else:
            count = len(comments)
            msg = f"Has {count} comment{'s' if count != 1 else ''}"
            self._announce(msg)

    def _get_comments_on_slide(self, window):
        """Get all comments on the current slide in a specific window.

        Args:
            window: DocumentWindow object

        Returns:
            list: Comment dictionaries
        """
        try:
            slide = window.View.Slide
            comments = []
            comment_count = slide.Comments.Count

            for i in range(1, comment_count + 1):
                try:
                    comment = slide.Comments.Item(i)
                    comments.append({
                        'text': comment.Text,
                        'author': comment.Author,
                        'datetime': comment.DateTime
                    })
                except Exception as e:
                    log.warning(f"Error reading comment {i} - {e}")

            return comments
        except Exception as e:
            log.debug(f"Could not get comments - {e}")
            return []
```

### 3.4 Advantages

- **Most reliable:** Uses the object that PowerPoint explicitly passed to the event
- **No timing issues:** The `sel` parameter always refers to the window that triggered the event
- **Clean code:** Straightforward object model navigation
- **No HWND matching:** Doesn't require Windows API calls

---

## 4. Solution 2: HWND Matching with Windows Collection

### 4.1 Concept

When `sel.Parent` is not available (rare edge cases), use Windows API to find which window has OS focus, then match it to the PowerPoint Windows collection.

### 4.2 Implementation

```python
import ctypes
from ctypes import windll

class PowerPointEventSink(COMObject):

    def WindowSelectionChange(self, sel):
        """Event handler with HWND fallback."""
        try:
            # Primary method: Use sel.Parent
            try:
                window = sel.Parent
                if window:
                    slide_index = window.View.Slide.SlideIndex
                    self._worker.on_slide_changed_event(window, slide_index)
                    return
            except:
                log.debug("sel.Parent failed, trying HWND matching")

            # Fallback: HWND matching
            window = self._get_focused_window_by_hwnd()
            if window:
                slide_index = window.View.Slide.SlideIndex
                self._worker.on_slide_changed_event(window, slide_index)

        except Exception as e:
            log.error(f"WindowSelectionChange error - {e}")

    def _get_focused_window_by_hwnd(self):
        """Find PowerPoint window with OS focus using HWND matching.

        Returns:
            DocumentWindow with focus, or None
        """
        try:
            # Get the foreground window (OS-focused window)
            foreground_hwnd = windll.user32.GetForegroundWindow()

            if not foreground_hwnd:
                return None

            # Iterate through all PowerPoint windows
            app = self._worker._ppt_app
            windows = app.Windows

            for i in range(1, windows.Count + 1):
                try:
                    window = windows.Item(i)
                    # Compare HWND
                    if window.HWND == foreground_hwnd:
                        log.info(f"Found focused window by HWND: {window.Presentation.Name}")
                        return window
                except:
                    continue

            # No match found
            log.warning("No PowerPoint window matched foreground HWND")
            return None

        except Exception as e:
            log.error(f"HWND matching failed - {e}")
            return None
```

### 4.3 When to Use HWND Matching

Use HWND matching when:
- `sel.Parent` returns None (shouldn't happen but defensive coding)
- Working with SlideShow windows (different object model)
- Need to verify which window has actual OS focus
- Debugging multi-window issues

### 4.4 Limitations

- HWND is 32-bit even on 64-bit systems (but this is guaranteed to work)
- Requires Windows API calls
- Slightly more complex code
- Theoretically, race condition if focus changes between event and HWND check

---

## 5. Solution 3: Iterate Windows Collection

### 5.1 Find Active Window Using Active Property

PowerPoint's `DocumentWindow.Active` property indicates if a window is active.

```python
def _get_active_window_from_collection(self):
    """Find the active window by iterating Windows collection.

    Returns:
        DocumentWindow that is active, or None
    """
    try:
        app = self._ppt_app
        windows = app.Windows

        for i in range(1, windows.Count + 1):
            try:
                window = windows.Item(i)
                if window.Active:
                    log.info(f"Found active window: {window.Presentation.Name}")
                    return window
            except:
                continue

        log.warning("No active window found in collection")
        return None

    except Exception as e:
        log.error(f"Failed to iterate Windows collection - {e}")
        return None
```

### 5.2 Combined Approach

```python
def WindowSelectionChange(self, sel):
    """Multi-strategy approach to finding the correct window."""
    window = None

    # Strategy 1: Use sel.Parent (most reliable)
    try:
        window = sel.Parent
        if window:
            log.debug("Using sel.Parent")
    except:
        pass

    # Strategy 2: HWND matching
    if not window:
        try:
            window = self._get_focused_window_by_hwnd()
            if window:
                log.debug("Using HWND matching")
        except:
            pass

    # Strategy 3: Iterate Windows.Active
    if not window:
        try:
            window = self._get_active_window_from_collection()
            if window:
                log.debug("Using Windows.Active iteration")
        except:
            pass

    # Strategy 4: Fallback to Application.ActiveWindow
    if not window:
        try:
            window = self._ppt_app.ActiveWindow
            log.debug("Fallback to Application.ActiveWindow")
        except:
            log.error("All window detection strategies failed")
            return

    # Process the slide change
    if window:
        try:
            slide_index = window.View.Slide.SlideIndex
            self._worker.on_slide_changed_event(window, slide_index)
        except Exception as e:
            log.error(f"Failed to process slide change - {e}")
```

---

## 6. Understanding Application.Windows Collection

### 6.1 Windows Collection Properties

```python
# Access all open document windows
windows = app.Windows
count = windows.Count  # Number of open windows

# Iterate windows
for i in range(1, count + 1):
    window = windows.Item(i)
    print(f"Window {i}: {window.Presentation.Name}")
    print(f"  Active: {window.Active}")
    print(f"  HWND: {window.HWND}")
    print(f"  View Type: {window.ViewType}")
```

### 6.2 Windows(1) Always Returns Active Window

**Important:** `Application.Windows(1)` always returns the active window, even though the collection is not reordered.

```python
# These should be equivalent (in theory)
active1 = app.ActiveWindow
active2 = app.Windows.Item(1)

# But Windows(1) is more reliable in practice
```

### 6.3 Per-Presentation Windows Collection

Each Presentation also has its own Windows collection:

```python
# Get windows for a specific presentation
presentation = app.Presentations.Item(1)
pres_windows = presentation.Windows

# A presentation can have multiple windows (like Excel)
for i in range(1, pres_windows.Count + 1):
    window = pres_windows.Item(i)
    print(f"Window {i} of {presentation.Name}")
```

### 6.4 Example: Listing All Windows

```python
def list_all_powerpoint_windows(app):
    """List all open PowerPoint windows with details.

    Args:
        app: PowerPoint.Application COM object
    """
    try:
        windows = app.Windows
        print(f"Total windows: {windows.Count}")

        for i in range(1, windows.Count + 1):
            try:
                window = windows.Item(i)
                print(f"\nWindow {i}:")
                print(f"  Presentation: {window.Presentation.Name}")
                print(f"  Active: {window.Active}")
                print(f"  HWND: {window.HWND}")
                print(f"  View Type: {window.ViewType}")
                print(f"  Window Number: {window.WindowNumber}")

                try:
                    slide = window.View.Slide
                    print(f"  Current Slide: {slide.SlideIndex}")
                except:
                    print(f"  Current Slide: N/A")

            except Exception as e:
                print(f"Error reading window {i}: {e}")

    except Exception as e:
        print(f"Failed to list windows: {e}")
```

---

## 7. Testing Strategy

### 7.1 Test Scenarios

1. **Single Presentation:**
   - Open one presentation
   - Navigate slides
   - Verify correct slide is detected

2. **Two Presentations - Sequential Access:**
   - Open `Pres1.pptx` and `Pres2.pptx`
   - Switch to Pres1, navigate to slide 5
   - Switch to Pres2, navigate to slide 3
   - Verify each window shows correct slides

3. **Two Presentations - Rapid Switching:**
   - Open two presentations
   - Rapidly switch between windows using Alt+Tab
   - Navigate slides in each
   - Verify no cross-contamination of slide data

4. **Multiple Windows of Same Presentation:**
   - Open one presentation
   - Create new window: View → New Window
   - Navigate to different slides in each window
   - Verify each window tracks independently

### 7.2 Debug Logging

```python
def WindowSelectionChange(self, sel):
    """Event handler with detailed logging."""
    try:
        # Log the selection object
        log.debug(f"WindowSelectionChange fired")
        log.debug(f"  sel object: {sel}")

        # Try to get window via sel.Parent
        try:
            window = sel.Parent
            log.debug(f"  sel.Parent: {window}")
            if window:
                log.debug(f"  Window.Presentation: {window.Presentation.Name}")
                log.debug(f"  Window.Active: {window.Active}")
                log.debug(f"  Window.HWND: {window.HWND}")
        except Exception as e:
            log.debug(f"  sel.Parent failed: {e}")

        # Compare with ActiveWindow
        try:
            active_window = self._worker._ppt_app.ActiveWindow
            log.debug(f"  ActiveWindow.Presentation: {active_window.Presentation.Name}")
            log.debug(f"  ActiveWindow.HWND: {active_window.HWND}")

            # Check if they match
            if window and window.HWND != active_window.HWND:
                log.warning("sel.Parent and ActiveWindow DIFFER!")
                log.warning(f"  sel.Parent points to: {window.Presentation.Name}")
                log.warning(f"  ActiveWindow points to: {active_window.Presentation.Name}")
        except Exception as e:
            log.debug(f"  ActiveWindow check failed: {e}")

    except Exception as e:
        log.error(f"WindowSelectionChange error: {e}")
```

### 7.3 Verification Checklist

- [ ] Single presentation works correctly
- [ ] Multiple presentations track independently
- [ ] Switching windows updates correctly
- [ ] Rapid switching doesn't cause errors
- [ ] Comments from correct presentation are announced
- [ ] No "wrong window" announcements
- [ ] Logs show correct window detection method used
- [ ] Performance is acceptable (no lag)

---

## 8. Recommended Final Implementation

### 8.1 Complete Event Sink

```python
class PowerPointEventSink(COMObject):
    """COM Event Sink for PowerPoint application events."""

    _com_interfaces_ = [EApplication, IDispatch]

    def __init__(self, worker):
        super().__init__()
        self._worker = worker
        self._last_slide_by_window = {}
        log.info("PowerPointEventSink initialized")

    def WindowSelectionChange(self, sel):
        """Called when selection changes in PowerPoint window.

        Handles multiple open presentations by using sel.Parent to get
        the specific window that triggered the event.

        Args:
            sel: Selection object (IDispatch)
        """
        try:
            log.debug("WindowSelectionChange event received")

            # Get the window that contains this selection
            window = self._get_window_from_selection(sel)

            if not window:
                log.warning("Could not determine window - skipping event")
                return

            # Get slide from the specific window
            try:
                slide = window.View.Slide
                slide_index = slide.SlideIndex
                presentation = window.Presentation

                # Create unique key for this window
                window_key = self._get_window_key(window)

                # Check if slide changed in THIS specific window
                last_slide = self._last_slide_by_window.get(window_key, -1)

                if slide_index != last_slide:
                    log.info(f"Slide changed: {presentation.Name} → slide {slide_index}")
                    self._last_slide_by_window[window_key] = slide_index
                    self._worker.on_slide_changed_event(window, slide_index)

            except Exception as e:
                log.debug(f"Could not get slide from window - {e}")

        except Exception as e:
            log.error(f"Error in WindowSelectionChange - {e}")

    def _get_window_from_selection(self, sel):
        """Get the DocumentWindow from a Selection object.

        Uses multiple strategies to reliably find the correct window.

        Args:
            sel: Selection object

        Returns:
            DocumentWindow or None
        """
        # Strategy 1: Use sel.Parent (most reliable)
        try:
            window = sel.Parent
            if window:
                log.debug(f"Window from sel.Parent: {window.Presentation.Name}")
                return window
        except Exception as e:
            log.debug(f"sel.Parent failed - {e}")

        # Strategy 2: Fallback to ActiveWindow
        try:
            window = self._worker._ppt_app.ActiveWindow
            log.debug(f"Fallback to ActiveWindow: {window.Presentation.Name}")
            return window
        except Exception as e:
            log.error(f"ActiveWindow also failed - {e}")
            return None

    def _get_window_key(self, window):
        """Create unique identifier for a window.

        Args:
            window: DocumentWindow object

        Returns:
            str: Unique key for this window
        """
        try:
            # Use presentation full name + HWND for uniqueness
            # (in case same file is opened multiple times)
            pres_name = window.Presentation.FullName
            hwnd = window.HWND
            return f"{pres_name}_{hwnd}"
        except:
            # Fallback to simple presentation name
            return window.Presentation.Name

    def SlideShowNextSlide(self, slideShowWindow):
        """Called when slide advances in slideshow mode.

        Args:
            slideShowWindow: SlideShowWindow object (IDispatch)
        """
        try:
            log.debug("SlideShowNextSlide event received")
            if self._worker and slideShowWindow:
                try:
                    slide_index = slideShowWindow.View.Slide.SlideIndex
                    log.info(f"Slideshow slide changed to {slide_index}")
                    # Note: SlideShowWindow doesn't have the same properties as DocumentWindow
                    # We can get the presentation via slideShowWindow.Presentation
                    self._worker.on_slideshow_slide_change(slideShowWindow, slide_index)
                except Exception as e:
                    log.debug(f"Could not get slideshow slide - {e}")
        except Exception as e:
            log.error(f"Error in SlideShowNextSlide - {e}")
```

### 8.2 Updated Worker

```python
class PowerPointWorker:
    """Background thread for PowerPoint COM operations."""

    def on_slide_changed_event(self, window, slide_index):
        """Called by event sink when slide changes.

        Args:
            window: DocumentWindow that triggered the event
            slide_index: New slide index (1-based)
        """
        log.info(f"Slide change event - slide {slide_index}")

        # Avoid duplicate announcements
        # (event sink already checks, but double-check)
        if slide_index == self._last_announced_slide:
            log.debug(f"Ignoring duplicate slide {slide_index}")
            return

        self._last_announced_slide = slide_index

        # Announce comments from the SPECIFIC window
        self._announce_slide_comments(window)

    def _announce_slide_comments(self, window):
        """Announce comment status for current slide in a specific window.

        Args:
            window: DocumentWindow to get comments from
        """
        comments = self._get_comments_on_slide(window)

        if not comments:
            self._announce("No comments")
            log.info("No comments on this slide")
        else:
            count = len(comments)
            msg = f"Has {count} comment{'s' if count != 1 else ''}"
            self._announce(msg)
            log.info(f"{msg}")

    def _get_comments_on_slide(self, window):
        """Get all comments on the current slide in a specific window.

        Args:
            window: DocumentWindow object

        Returns:
            list: Comment dictionaries
        """
        try:
            slide = window.View.Slide
            comments = []
            comment_count = slide.Comments.Count
            log.debug(f"Found {comment_count} comments on slide")

            for i in range(1, comment_count + 1):
                try:
                    comment = slide.Comments.Item(i)
                    comments.append({
                        'text': comment.Text,
                        'author': comment.Author,
                        'datetime': comment.DateTime
                    })
                except Exception as e:
                    log.warning(f"Error reading comment {i} - {e}")

            return comments
        except Exception as e:
            log.debug(f"Could not get comments - {e}")
            return []
```

---

## 9. References and Sources

### Microsoft Documentation
- [DocumentWindow object (PowerPoint)](https://msdn.microsoft.com/en-us/vba/powerpoint-vba/articles/documentwindow-object-powerpoint)
- [DocumentWindows object (PowerPoint)](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.documentwindows)
- [Application.Windows property (PowerPoint)](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.Application.Windows)
- [DocumentWindow.Selection property (PowerPoint)](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.documentwindow.selection)
- [Selection.Parent property (PowerPoint)](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.selection.parent)
- [DocumentWindow.Presentation Property](https://learn.microsoft.com/en-us/previous-versions/office/office-12/ff760491(v=office.12))
- [Application.ActiveWindow property (PowerPoint)](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.activewindow)
- [Application.WindowSelectionChange event (PowerPoint)](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.Application.WindowSelectionChange)

### NVDA Source Code
- [NVDA GitHub Repository](https://github.com/nvaccess/nvda)
- [NVDA appModules Source](https://github.com/nvaccess/nvda/tree/master/source/appModules)
- [NVDA 2025.3.2 Developer Guide](https://download.nvaccess.org/documentation/developerGuide.html)

### Community Resources
- [PowerPoint VBA Events - VBA Express Forum](http://www.vbaexpress.com/forum/archive/index.php/t-13572.html)
- [PowerPoint Application Events in VBA - YOUpresent](http://youpresent.co.uk/powerpoint-application-events-in-vba/)
- [Events: Mastering PowerPoint VBA Events - FasterCapital](https://fastercapital.com/content/Events--Mastering-PowerPoint-VBA-Events-for-Interactive-Presentations.html)
- [PowerPoint HWND Discussion - MSDN Forums](https://social.msdn.microsoft.com/Forums/officeapps/en-US/e9463be7-23de-429b-a571-1cd2414c772a/powerpoint-is-documentwindowhwnd-an-int-valid-on-64bit)

---

## Document Information

- **Created:** December 10, 2025
- **Author:** Research and Analysis
- **Version:** 1.0
- **Purpose:** Solve multi-window focus detection in PowerPoint NVDA addon
- **Status:** Complete - ready for implementation
