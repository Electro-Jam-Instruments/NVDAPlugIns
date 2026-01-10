# Failed Approaches Log

Document of failed implementation attempts to prevent repeating mistakes.

## Slideshow "Has Notes" Prefix (v0.0.61-v0.0.67)

**Goal:** Announce "has notes" BEFORE slide number/title during slideshow.

### Attempt 1: `_get_name()` Override (FAILED)
```python
class CustomSlideShowWindow(SlideShowWindow):
    def _get_name(self):
        if self.appModule._worker._has_meeting_notes():
            return f"has notes, {super()._get_name()}"
        return super()._get_name()
```
**Result:** `_get_name()` was NEVER called - slideshow uses `handleSlideChange()` → `reportFocus()` path, not normal focus flow.

### Attempt 2: `reportFocus()` Override (PARTIAL)
```python
def reportFocus(self):
    has_notes = self.appModule._worker._has_meeting_notes()
    if has_notes:
        ui.message(f"has notes, {super()._get_name()}")
    else:
        super().reportFocus()
```
**Result:** `reportFocus()` WAS called, but `has_notes` always returned False due to threading race condition.

### Attempt 3: Direct `self.currentSlide.NotesPage` (FAILED)
**Result:** `AttributeError: 'Slide' object has no attribute 'NotesPage'`
**Why:** `self.currentSlide` is an NVDA wrapper object, not a PowerPoint COM object.

### Attempt 4: Access via `self.currentSlide.Parent` (FAILED)
**Result:** `AttributeError: 'Slide' object has no attribute 'Parent'`
**Why:** The NVDA Slide wrapper doesn't have COM object Parent property.

### Attempt 5: Type diagnostics (DISCOVERY)
**Finding:** `self.currentSlide` has NVDA properties (`appModule`, `TextInfo`) not COM properties (`NotesPage`, `SlideIndex`). Cannot access PowerPoint COM object directly from overlay class.

### Key Learning
`CustomSlideShowWindow.currentSlide` is an NVDA wrapper, NOT a PowerPoint COM object. Must access slide data through worker thread's COM connection, not through the overlay class.

---

## Focus-Based Voice Typing Detection (v0.0.1)

**Goal:** Detect Windows Voice Typing window to silence NVDA.

### Attempt: Focus change events
```python
def event_foregroundChange(self, obj, nextHandler):
    if "Voice Typing" in obj.name:
        self._silence_speech()
```
**Result:** Events never fired.
**Why:** Voice Typing is a lightweight overlay that doesn't take traditional Windows focus.

---

## Timer-Based Voice Typing Window Polling (v0.0.2)

**Goal:** Poll for Voice Typing window presence.
```python
def _poll_for_voice_typing():
    while True:
        if find_window("Voice Typing"):
            silence_speech()
        sleep(0.5)
```
**Result:** Worked but rejected.
**Why:** User requirement: "I'm not a fan of a timer" - no polling approach acceptable.

---

## Direct GetActiveObject for COM (v0.0.11-v0.0.12)

**Goal:** Connect to PowerPoint COM object.
```python
from comtypes.client import GetActiveObject
ppt = GetActiveObject("PowerPoint.Application")
```
**Result:** `WinError -2147221021 Operation unavailable`
**Why:** NVDA runs with UIAccess privileges, blocking direct COM access to lower-privilege processes.

---

## Type Library Loading (v0.0.16-v0.0.20)

**Goal:** Load PowerPoint type library for COM interfaces.
```python
from comtypes.client import GetModule
GetModule(['{91493440-5A91-11CF-8700-00AA0060263B}', 1, 0])
```
**Result:** `[WinError -2147319779] Library not registered`
**Why:** PowerPoint's type library isn't reliably registered, especially on Office 365.

---

## event_NVDAObject_init for UIA Objects (v0.0.36)

**Goal:** Modify comment card names at object initialization.
```python
def event_NVDAObject_init(self, obj):
    if obj.UIAAutomationId.startswith('cardRoot_'):
        obj.name = self._reformat_comment(obj.name)
```
**Result:** Event never fired for comment cards.
**Why:** `event_NVDAObject_init` is NOT reliably called for UIA objects in PowerPoint. Properties may not be available at init time.

---

## pywin32 for COM Automation (Various)

**Goal:** Use win32com.client for PowerPoint automation.
**Result:** Silent failures, DLL conflicts.
**Why:**
- NVDA uses comtypes internally
- pywin32 DLLs conflict with NVDA process
- pywin32 isn't in NVDA runtime

---

## Explicit Import Alias for Inheritance (v0.0.4-v0.0.8)

**Goal:** Inherit from built-in AppModule.
```python
from nvdaBuiltin.appModules.powerpnt import AppModule as BuiltinAppModule
class AppModule(BuiltinAppModule):
    pass
```
**Result:** Module did NOT load. Silent failure.
**Why:** Unknown - appears logically equivalent to working pattern but doesn't work in NVDA's module loading system.

---

## reportNewSlide() on SlideShowWindow (v0.0.76)

**Goal:** Suppress slideshow content reading.
```python
class CustomSlideShowWindow(SlideShowWindow):
    def reportNewSlide(self):
        pass  # Override here
```
**Result:** No effect - full slide content still read.
**Why:** `reportNewSlide()` is a method on `ReviewableSlideshowTreeInterceptor`, NOT `SlideShowWindow`. Wrong class.

---

## Complete reportNewSlide() Suppression (v0.0.77)

**Goal:** Completely silence slideshow announcements.
```python
class CustomTreeInterceptor(ReviewableSlideshowTreeInterceptor):
    def reportNewSlide(self):
        pass  # Do nothing
```
**Result:** Complete silence - nothing announced.
**Why:** Parent's `reportNewSlide()` handles ALL slide announcements. Must call it with modified behavior, not suppress entirely.

---

## Key Patterns to Avoid

| What | Why It Fails |
|------|--------------|
| `GetActiveObject()` | UIAccess privilege blocking |
| `GetModule()` for type library | Type library not registered |
| `event_NVDAObject_init` for UIA | Not called reliably |
| `self.currentSlide` for COM access | It's an NVDA wrapper, not COM |
| pywin32 | DLL conflicts with NVDA |
| Explicit alias inheritance | Silent module loading failure |
| Override on wrong class | No effect |
| Complete method suppression | Removes all functionality |
