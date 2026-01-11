# Pitfalls to Avoid

Critical failures from development that looked like they should work but didn't.

## Pitfall 1: Explicit Import Alias for Inheritance

**What we tried (v0.0.4-v0.0.8):**
```python
from nvdaBuiltin.appModules.powerpnt import AppModule as BuiltinAppModule
class AppModule(BuiltinAppModule):
    pass
```

**What happened:** Module did NOT load. No log entries, no errors, just silent failure.

**Why it failed:** Unknown - appears logically equivalent to the working pattern but doesn't work in NVDA's module loading system.

**The fix:** Use exact NVDA doc pattern:
```python
from nvdaBuiltin.appModules.powerpnt import *
class AppModule(AppModule):
    pass
```

---

## Pitfall 2: Calling super() on Optional Event Hooks

**What we tried (v0.0.10):**
```python
def event_appModule_gainFocus(self):
    super().event_appModule_gainFocus()  # CRASH
```

**What happened:** `AttributeError: 'super' object has no attribute 'event_appModule_gainFocus'`

**Why it failed:** `event_appModule_gainFocus` is an optional hook - parent class doesn't define it.

**The fix:** Don't call super() on optional hooks:
```python
def event_appModule_gainFocus(self):
    # No super() call needed - delegate to worker thread
    if self._worker:
        self._worker.request_initialize()
```

---

## Pitfall 3: Direct COM Access with GetActiveObject

**What we tried (v0.0.11-v0.0.12):**
```python
from comtypes.client import GetActiveObject
ppt = GetActiveObject("PowerPoint.Application")
```

**What happened:** `WinError -2147221021 Operation unavailable`

**Why it failed:** NVDA runs with UIAccess privileges. Windows blocks high-privilege processes from directly accessing COM in lower-privilege processes.

**The fix:** Use NVDA's comHelper:
```python
import comHelper
ppt = comHelper.getActiveObject("PowerPoint.Application", dynamic=True)
```

---

## Pitfall 4: Loading PowerPoint Type Library

**What we tried (v0.0.16-v0.0.20):**
```python
from comtypes.client import GetModule
GetModule(['{91493440-5A91-11CF-8700-00AA0060263B}', 1, 0])
```

**What happened:** `[WinError -2147319779] Library not registered`

**Why it failed:** PowerPoint's type library isn't reliably registered on all systems, especially Office 365 installations.

**The fix:** Define the EApplication interface locally (copy NVDA's pattern):
```python
class EApplication(IDispatch):
    _iid_ = GUID("{914934C2-5A91-11CF-8700-00AA0060263B}")
    _methods_ = []
    _disp_methods_ = [...]
```

---

## Pitfall 5: reportNewSlide() on Wrong Class

**What we tried (v0.0.76):**
```python
class CustomSlideShowWindow(SlideShowWindow):
    def reportNewSlide(self):  # Override here
        pass
```

**What happened:** No effect - full slide content still read aloud.

**Why it failed:** `reportNewSlide()` is a method on `ReviewableSlideshowTreeInterceptor`, NOT on `SlideShowWindow`. Overriding on the wrong class does nothing.

**The fix:** Override on TreeInterceptor and link via `treeInterceptorClass`:
```python
class CustomTreeInterceptor(ReviewableSlideshowTreeInterceptor):
    def reportNewSlide(self):
        # Our override here

class CustomSlideShowWindow(SlideShowWindow):
    treeInterceptorClass = CustomTreeInterceptor
```

---

## Pitfall 6: Suppressing reportNewSlide() Entirely

**What we tried (v0.0.77):**
```python
class CustomTreeInterceptor(ReviewableSlideshowTreeInterceptor):
    def reportNewSlide(self):
        pass  # Do nothing
```

**What happened:** Complete silence - nothing announced on slide change.

**Why it failed:** The parent's `reportNewSlide()` handles all slide announcements. Suppressing it entirely means no announcement at all.

**The fix:** Speak the first line (title) instead of sayAll:
```python
def reportNewSlide(self, suppressSayAll=False):
    info = self.selection
    info.expand(textInfos.UNIT_LINE)
    speech.speakTextInfo(info, ...)
```

---

## Pitfall 7: pywin32 Instead of comtypes

**What happened:** Early attempts used `win32com.client` - fails silently or with DLL conflicts.

**Why it failed:**
- NVDA uses comtypes internally
- pywin32 DLLs conflict with NVDA's process
- pywin32 isn't in NVDA runtime

**The fix:** Always use comtypes for COM in NVDA addons.

---

## Pitfall 8: Heavy Work in Event Handlers

**What we tried (v0.0.9):**
```python
def event_appModule_gainFocus(self):
    ppt = comHelper.getActiveObject(...)  # Blocks here
    comments = ppt.ActiveWindow.View.Slide.Comments
```

**What happened:** NVDA stopped speaking during focus change.

**Why it failed:** Event handlers that block prevent NVDA from processing speech queue.

**The fix (v0.0.14+):** Delegate to worker thread (no arbitrary delays):
```python
def event_appModule_gainFocus(self):
    # Non-blocking - just signals worker thread
    if self._worker:
        self._worker.request_initialize()
```

> **Note:** Early versions used `core.callLater(100, ...)` which is deprecated. Worker thread is preferred - no arbitrary delays, proper lifecycle management.

---

## Pitfall 9: Worker Thread Data Stale at Event Time

**What we tried:** Read cached worker data in `event_NVDAObject_init`.

**What happened:** Stale/missing data - worker hadn't finished querying yet.

**Why it failed:** Main thread events fire immediately. Worker thread COM queries take time.

**The fix:** Use lazy `_get_name()` in overlay class - called later when NVDA needs the name.

---

## Pitfall 10: Non-Breaking Spaces in PowerPoint Text (v0.0.42)

**What we tried:**
```python
if "started by" in comment_name:
    author = comment_name.split("started by")[1]
```

**What happened:** Parsing silently failed - author extraction returned wrong results.

**Why it failed:** PowerPoint uses U+00A0 (non-breaking space) not U+0020 (regular space). String operations like `.split()` and `in` don't match.

**The fix:** Normalize all whitespace first:
```python
import re
# Convert all whitespace (including U+00A0) to regular spaces
name_normalized = re.sub(r'\s+', ' ', comment_name)
if "started by" in name_normalized:
    author = name_normalized.split("started by")[1]
```

**Where this matters:**
- Comment name parsing
- Any string matching on PowerPoint-generated text

---

## Pitfall 11: SlideShowNextSlide Doesn't Fire for First Slide

**What we tried:** Rely solely on `SlideShowNextSlide` event for slideshow slide data.

**What happened:** First slide had no prefix (notes/comments not announced).

**Why it failed:** PowerPoint's `SlideShowNextSlide` event does NOT fire when slideshow starts - only on subsequent slides.

**The fix:** Cache first slide data in `SlideShowBegin` event:
```python
def on_slideshow_begin(self, wn):
    self._in_slideshow = True
    # CRITICAL: Cache first slide immediately
    self._cache_slideshow_slide_data(wn)
```

---

## Pitfall 12: Using ActiveWindow with Multiple Presentations

**What we tried:**
```python
def WindowSelectionChange(self, sel):
    slide = self._ppt_app.ActiveWindow.View.Slide  # WRONG
```

**What happened:** With two presentations open, comments from the wrong presentation were announced.

**Why it failed:** `Application.ActiveWindow` returns the last-used window in the application, not necessarily the window that triggered the `WindowSelectionChange` event. During rapid window switching, `ActiveWindow` can lag behind the actual focus.

**Real-world impact:**
1. User switches to `Presentation2.pptx`
2. `WindowSelectionChange` event fires
3. Plugin checks `ActiveWindow.View.Slide` → still pointing to Presentation1
4. Plugin announces comments from the WRONG presentation

**The fix:** Use `sel.Parent` to get the correct window:
```python
def WindowSelectionChange(self, sel):
    # sel.Parent is the DocumentWindow that triggered the event
    window = sel.Parent
    slide = window.View.Slide  # CORRECT - always the right window
```

**Key insight:** The `sel` parameter passed to `WindowSelectionChange` always refers to the selection in the window that triggered the event. Its `Parent` property returns that specific `DocumentWindow`.

---

## Pitfall 13: self.currentSlide is NVDA Wrapper, Not COM Object

**What we tried (v0.0.63):**
```python
class CustomSlideShowWindow(SlideShowWindow):
    def _check_slide_has_notes(self):
        notes_page = self.currentSlide.NotesPage  # Access COM property
```

**What happened:** `AttributeError: 'Slide' object has no attribute 'NotesPage'`

**Why it failed:** In overlay classes, `self.currentSlide` is an NVDA wrapper object, NOT a PowerPoint COM object. It has NVDA properties (`appModule`, `TextInfo`, `APIClass`) not COM properties (`NotesPage`, `SlideIndex`).

**The fix:** Access COM data via worker thread which has the actual COM connection:
```python
# Worker thread has the COM object
slide = self._slideshow_window.View.Slide  # Actual COM object
notes_page = slide.NotesPage  # Works
```

---

## Pitfall 14: _get_name() Not Called in All Contexts

**What we tried (v0.0.61):**
```python
class CustomSlideShowWindow(SlideShowWindow):
    def _get_name(self):
        # Prepend notes info
        return f"has notes, {super()._get_name()}"
```

**What happened:** `_get_name()` was NEVER called during slideshow mode.

**Why it failed:** In slideshow mode, NVDA uses `handleSlideChange()` → `reportFocus()` flow instead of the normal focus event flow that calls `_get_name()`.

**The fix:** Override `reportFocus()` for slideshow, or use TreeInterceptor's `reportNewSlide()`:
```python
def reportFocus(self):
    # Called in slideshow - can prepend info here
    if has_notes:
        ui.message("has notes")
    super().reportFocus()
```

---

## Quick Reference: What Works

| Task | Wrong Approach | Right Approach |
|------|----------------|----------------|
| Inherit built-in | Explicit alias | `import *` then `class X(X)` |
| COM access | `GetActiveObject()` | `comHelper.getActiveObject()` |
| COM events | Load type library | Define interface locally |
| Event handler super() | Always call | Only if parent has method |
| Heavy work | In event handler | Delegate to worker thread |
| Slideshow override | On SlideShowWindow | On TreeInterceptor |
| Fresh data | `event_NVDAObject_init` | Lazy `_get_name()` |
| PPT string parsing | Direct string ops | Normalize whitespace first |
| First slideshow slide | SlideShowNextSlide | Cache in SlideShowBegin |
| Multi-window slide access | `ActiveWindow.View.Slide` | `sel.Parent.View.Slide` |
| Overlay COM access | `self.currentSlide.NotesPage` | Worker thread COM object |
| Slideshow name override | `_get_name()` | `reportFocus()` or TreeInterceptor |
