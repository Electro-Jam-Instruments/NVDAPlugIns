# PowerPoint Automation - Decisions

## Decision Log

### 1. COM for Data, UIA for Focus

**Decision:** Use COM API to read comment data, UIA to manage focus
**Date:** December 2025
**Status:** Final

**Rationale:**
- COM provides reliable access to `Slide.Comments` collection
- Comments pane is UIA-enabled (NetUIHWNDElement)
- UIA focus more reliable than COM selection for UI elements
- Best of both worlds

**Research:**
- `research/PowerPoint_Comment_Focus_Navigation_Research.md`
- `research/PowerPoint-COM-Automation-Research.md`

---

### 2. View Management Strategy

**Decision:** Auto-switch to Normal view when needed
**Date:** December 2025
**Status:** Final

**Rationale:**
- Comments pane only accessible in Normal view
- User should not have to manually switch
- `ActiveWindow.ViewType = 9` is reliable

**ViewType Constants:**
- Normal = 9
- Slide Sorter = 5
- Notes = 10
- Outline = 6
- Slide Master = 3
- Reading = 50

**Research:** `research/PowerPoint_Comment_Focus_Navigation_Research.md`

---

### 3. @Mention Detection via Regex

**Decision:** Parse @mentions from comment text using regex
**Date:** December 2025
**Status:** Final

**Rationale:**
- No structured mention data in COM API
- @mentions stored as plain text in `Comment.Text`
- Regex with fuzzy matching handles variations

**Pattern:** `@([A-Za-z\u00C0-\u024F][\w\u00C0-\u024F]*(?:[-'][\w\u00C0-\u024F]+)?(?:\s+[A-Za-z\u00C0-\u024F][\w\u00C0-\u024F]*)*)`

**Research:** `research/powerpoint_mention_detection_research.md`

---

### 4. Open Comments Pane via ExecuteMso

**Decision:** Use `CommandBars.ExecuteMso("ReviewShowComments")` to open pane
**Date:** December 2025
**Status:** Final

**Rationale:**
- Most reliable way to ensure pane is visible
- Try multiple command names for Office version compatibility
- Fallback commands: "ShowComments", "CommentsPane"

**Research:** `research/PowerPoint_Comment_Focus_Navigation_Research.md`

---

### 5. Slide Change Detection via Focus Events

**Decision:** Track slide changes via NVDA focus events + COM polling
**Date:** December 2025
**Status:** Final

**Rationale:**
- NVDA's `event_gainFocus` fires on selection changes
- Check `ActiveWindow.View.Slide.SlideIndex` for actual slide
- Compare to cached index to detect changes

**Research:** `research/NVDA_PowerPoint_Native_Support_Analysis.md`

---

## Backlogged Decisions

### Threaded Comment Replies

**Issue:** Modern comments support reply threads
**Status:** Backlogged for post-MVP

**Current Approach:** Treat all comments as flat list
**Future:** Navigate parent/child hierarchy
