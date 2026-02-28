# Architecture Decisions

Key technical decisions made during development, with rationale.

## Decision 1: AppModule Inheritance Pattern

**Choice:** `from nvdaBuiltin.appModules.powerpnt import *` then `class AppModule(AppModule)`

**Why:** Only pattern that works. Explicit alias breaks module loading. Base class loses built-in features.

**Verified:** v0.0.9+

**Details:** See [NVDA Addon Development Guide](../guides/nvda-addon-development.md)

---

## Decision 2: COM Access Method

**Choice:** `comHelper.getActiveObject()` not direct `GetActiveObject()`

**Constraint (Windows):** NVDA runs with UIAccess privileges. Windows security blocks high-privilege processes from directly accessing COM in lower-privilege processes (PowerPoint). This causes `WinError -2147221021`.

**Our choice:** Adopt NVDA's `comHelper` module, which handles the privilege bridging correctly.

**Verified:** v0.0.13+

---

## Decision 3: Use comtypes, Not pywin32

**Choice:** Use `comtypes` for all COM automation

**Why:**
- NVDA uses comtypes internally
- pywin32 DLLs conflict with NVDA process
- comtypes already in NVDA runtime

---

## Decision 4: Define EApplication Interface Locally

**Choice:** Define PowerPoint's COM event interface in our code, not from type library

**Why:** Type library loading fails with "Library not registered" in many environments. Local definition is what NVDA's built-in PowerPoint module does.

**Verified:** v0.0.21+

**Key GUID:** `{914934C2-5A91-11CF-8700-00AA0060263B}` (EApplication events)

---

## Decision 5: Worker Thread for COM Operations

**Choice:** Dedicated background thread with work queue

**Why:**
- NVDA maintainers recommend threads for continuous/repeated work
- Worker thread allows proper COM initialization (STA) and cleanup
- Clean lifecycle management (start/stop with app focus)
- No arbitrary delays - work executes as soon as thread processes it

**Verified:** v0.0.14+

---

## Decision 6: Event-Driven Slide Detection

**Choice:** Use COM events (WindowSelectionChange, SlideShowNextSlide) instead of polling

**Why:**
- Instant response vs polling latency
- No CPU waste
- Proven pattern - NVDA's built-in PowerPoint module uses this

---

## Decision 7: Overlay Class with Lazy _get_name()

**Choice:** Use overlay class with `_get_name()` property for dynamic name modification

**Why:** Worker thread data may be stale when NVDA's `event_NVDAObject_init` fires. Lazy evaluation via `_get_name()` queries fresh data when NVDA actually needs the name.

**Verified:** v0.0.76+

---

## Decision 8: Override reportNewSlide on TreeInterceptor

**Choice:** Override `reportNewSlide()` on `ReviewableSlideshowTreeInterceptor`, not `SlideShowWindow`

**Why:** In slideshow mode, `reportNewSlide()` is a method on the TreeInterceptor class, not on the SlideShowWindow overlay. Overriding on the wrong class has no effect.

**Verified:** v0.0.78

**Details:** See [Slideshow Override Reference](../reference/slideshow-override.md)

---

## Decision 9: Target Modern Comments Only

**Choice:** Only support Modern Comments (PowerPoint 365), not legacy comments

**Why:**
- Legacy comments use different COM API
- Modern comments are the standard going forward
- Simplifies implementation
- Target users are on 365

---

## Decision 10: No Comment Resolution Status

**Choice:** Don't try to detect resolved vs unresolved comments (deferred)

**Why:**
- Resolved status NOT exposed in PowerPoint's COM/VBA API
- Unlike Word, PowerPoint's `Comment` object has no `Done` or `Resolved` property
- The status exists ONLY in OOXML (`status="active|resolved|closed"` in XML)

**Why OOXML approach doesn't work:**
1. **File is locked** - PowerPoint holds exclusive lock on open .pptx files
2. **Shadow copy attempt** - We tried using Windows VSS to read a shadow copy, but this is:
   - Complex to implement correctly
   - Unreliable across Windows versions
   - Requires elevated permissions in some cases
3. **Temp copy approach** - Copying the file while open sometimes works, sometimes doesn't depending on Windows caching

**Why UIA approach doesn't work:**
- Unlike Word, Microsoft has NOT documented PowerPoint-specific UIA custom properties for comment resolution
- The Comments pane checkbox is NOT reliably exposed via UIA ToggleState
- Would require Comments pane to be open (users may have it closed)

**Approaches investigated (Dec 2025):**
1. COM API - No property exists
2. OOXML direct read - File locked
3. OOXML shadow copy - Complex and fragile
4. OOXML temp copy - Unreliable
5. UIA toggle state - Not exposed

**Future:** Revisit if Microsoft exposes resolution status in COM API

---

## Decision 11: Don't Block Event Handlers

**Choice:** Delegate heavy work to worker thread from event handlers

**Why:** Event handlers that block prevent NVDA from speaking.

**Current pattern (v0.0.14+):**
```python
def event_appModule_gainFocus(self):
    # Non-blocking - just sets a flag for worker thread
    if self._worker:
        self._worker.request_initialize()
```

**Deprecated pattern (v0.0.11-v0.0.13):**
```python
# DON'T USE - arbitrary 100ms delay is fragile
def event_appModule_gainFocus(self):
    core.callLater(100, self._deferred_initialization)
```

Worker thread is preferred because it has no arbitrary delays and provides proper lifecycle management.

---

## Decision 12: COM for Data, UIA for Focus

**Choice:** Use COM API to read comment data, UIA to manage Comments pane focus

**Constraint (NVDA):** NVDA blocks UIA for PowerPoint's main content classes (`paneClassDC`, `mdiClass`) because Microsoft's UIA implementation was incomplete (NVDA Issue #3578). This may change as UIA improves.

**Our choice given the constraint:**
- **COM** for slide content, comments, and notes (works regardless of NVDA's UIA stance)
- **UIA** for focusing task panes like Comments (uses `NetUIHWND`, which NVDA supports)

This hybrid approach works now and would simplify if NVDA later enables UIA for more PowerPoint classes.

---

## Decision 13: Auto-Switch to Normal View

**Choice:** Auto-switch to Normal view when accessing comments

**Why:**
- Comments pane only accessible in Normal view
- User should not have to manually switch
- `ActiveWindow.ViewType = 9` is reliable

**ViewType Constants:**
- Normal = 9, Slide Sorter = 5, Notes = 10, Outline = 6, Slide Master = 3, Reading = 50

---

## Decision 14: @Mention Detection via Regex

**Choice:** Parse @mentions from comment text using regex pattern

**Why:**
- No structured mention data in COM API
- @mentions stored as plain text in `Comment.Text`
- Regex with fuzzy matching handles name variations

---

## Decision 15: Open Comments Pane via ExecuteMso

**Choice:** Use `CommandBars.ExecuteMso("ReviewShowComments")` to open pane

**Why:**
- Most reliable way to ensure pane is visible
- Try multiple command names for Office version compatibility
- Fallback commands: "ShowComments", "CommentsPane"

---

## Decision 16: Flat Comment List (Deferred Threading)

**Choice:** Treat all comments as flat list, defer thread reply navigation

**Why:**
- Modern comments support reply threads but flat list is simpler
- Thread navigation is post-MVP enhancement
