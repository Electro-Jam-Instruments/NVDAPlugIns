# Windows Accessibility - Decisions

## Decision Log

### 1. UIA for Comments Pane Focus

**Decision:** Use UI Automation to focus comments in the task pane
**Date:** December 2025
**Status:** Final

**Rationale:**
- Comments pane uses NetUIHWNDElement (UIA-enabled)
- NVDA disables UIA for main PowerPoint content (paneClassDC, mdiClass)
- But task panes remain UIA-accessible
- `IUIAutomation.SetFocus()` reliably moves keyboard focus

**Implementation:**
```python
automation = CreateObject("{ff48dba4-60ef-4201-aa87-54103eef594e}")
element.SetFocus()
```

**Research:** `research/PowerPoint-UIA-Research.md`

---

### 2. Window Class Targeting

**Decision:** Target specific window classes for UIA operations
**Date:** December 2025
**Status:** Final

**Key Classes:**
| Class | UIA Status | Purpose |
|-------|------------|---------|
| paneClassDC | Disabled by NVDA | Main content |
| mdiClass | Disabled by NVDA | MDI container |
| NetUIHWND | Enabled | Ribbon, task panes |
| screenClass | Disabled by NVDA | Slideshow |

**Rationale:** Must understand which windows support UIA

**Research:** `research/NVDA_PowerPoint_Native_Support_Analysis.md`

---

### 3. User Identity Detection

**Decision:** Use Windows SecurLib for display name, fallback to environment
**Date:** December 2025
**Status:** Final

**Rationale:**
- `GetUserNameExW(NameDisplay)` returns user's display name
- Matches what PowerPoint uses for @mentions
- Fallback to `%USERNAME%` if SecurLib fails

**Research:** `research/powerpoint_mention_detection_research.md`

---

## Notes

### NVDA's PowerPoint UIA Philosophy

NVDA deliberately disables UIA for PowerPoint content windows because Microsoft's UIA implementation is incomplete. From NVDA Issue #3578:

> "Microsoft has now tried to provide an accessibility implementation for PowerPoint using UI Automation. But as usual its far from complete, yet at the same time cripples any existing support/hacks by other ATs."

Our approach: Use COM for content, UIA only for task pane focus.
