# NVDA PowerPoint Community Add-ons Research

## Executive Summary

This research document catalogs existing NVDA plugins and add-ons that enhance PowerPoint accessibility. The investigation reveals that **no dedicated PowerPoint accessibility add-on currently exists** in the NVDA community. PowerPoint support is primarily built into NVDA's core functionality through the `powerpnt.py` app module, with ongoing development addressing known limitations.

**Key Finding:** The gap in PowerPoint-specific add-ons presents a significant opportunity for the A11Y PowerPoint NVDA Plug-In project to address unmet user needs, particularly around comment navigation, shape accessibility, and AI-powered image descriptions.

---

## Table of Contents

1. [NVDA Add-on Store Search Results](#1-nvda-add-on-store-search-results)
2. [GitHub Search Results](#2-github-search-results)
3. [NVDA Core PowerPoint Support Analysis](#3-nvda-core-powerpoint-support-analysis)
4. [Related Office Add-ons Analysis](#4-related-office-add-ons-analysis)
5. [Known PowerPoint Accessibility Issues](#5-known-powerpoint-accessibility-issues)
6. [Feature Comparison: NVDA vs JAWS for PowerPoint](#6-feature-comparison-nvda-vs-jaws-for-powerpoint)
7. [Gap Analysis](#7-gap-analysis)
8. [Recommendations](#8-recommendations)
9. [Contacts and Resources](#9-contacts-and-resources)
10. [References](#10-references)

---

## 1. NVDA Add-on Store Search Results

### Sources Searched
- NVDA Add-on Store (https://nvda.store/)
- NVDA Add-ons Directory (https://nvda-addons.org/)
- Legacy NVDA Community Add-ons (https://addons.nvda-project.org/)

### PowerPoint-Specific Add-ons Found

**Result: NONE**

No dedicated PowerPoint accessibility add-ons exist in any of the NVDA add-on repositories. The search covered:
- "PowerPoint" - No results
- "Presentation" - No results
- "Slides" - No results
- "PPT" - No results

### Office-Related Add-ons Found

| Add-on Name | Version | Status | PowerPoint Support | Notes |
|-------------|---------|--------|-------------------|-------|
| Office Desk | 25.09 | **END OF LIFE** (Sept 2025) | No | Word only |
| wordAccessEnhancement | 3.7.2 | Active | No | Word only |
| Outlook Extended | 3.2 | Active | No | Outlook only |

**Conclusion:** The NVDA community has not developed any dedicated PowerPoint accessibility add-ons. All PowerPoint support comes from NVDA core.

---

## 2. GitHub Search Results

### Repositories Searched
- nvaccess/nvda (NVDA core repository)
- nvdaaddons organization
- General GitHub search: "NVDA PowerPoint"

### Key Findings

#### NVDA Core PowerPoint App Module (`powerpnt.py`)
- **Location:** nvaccess/nvda/source/appModules/powerpnt.py
- **Implementation:** COM-based (not UI Automation)
- **Last Major Update:** Ongoing improvements through 2024

#### Relevant Pull Requests (Recent)

| PR Number | Title | Status | Description |
|-----------|-------|--------|-------------|
| #17015 | Fix PowerPoint caret reporting for wide characters | Merged | Enhanced text handling in PowerPoint |
| #17004 | Enable cursor movement with braille display routing keys | Merged | Added braille display support for PowerPoint |
| #7046 | Added chart support for Word and PowerPoint | Merged | Extended chart accessibility beyond Excel |

#### Open Issues Related to PowerPoint

| Issue | Title | Status | Impact |
|-------|-------|--------|--------|
| #16161 | NVDA doesn't perceive some objects and groups in slideshow | Open | Grouped objects and alt text ignored |
| #12719 | NVDA not reading PowerPoint Outline view | Open | "Linefeed" only, no actual text |
| #9941 | Report more accurate position of text with NVDA+delete | Open | Position reporting inaccurate |

#### Historical Issues (Resolved)

| Issue | Title | Resolution | Version |
|-------|-------|------------|---------|
| #4850 | PowerPoint slideshow reading | Fixed | NVDA 2015.2 |
| #3578 | PowerPoint 2013 text box announcement | Fixed | Added to badUIAWindowClasses |
| #9101 | Cursor routing from braille display | Fixed | PR #17004 |

---

## 3. NVDA Core PowerPoint Support Analysis

### Implementation Approach

The NVDA PowerPoint app module uses **COM (Component Object Model)** to interact with PowerPoint's object model, rather than UI Automation. This approach provides:

- Direct access to presentation structure
- Slide content and properties
- Shapes, images, and text frames
- Speaker notes and comments

### Key Classes in powerpnt.py

| Class | Purpose | Key Features |
|-------|---------|--------------|
| `PaneClassDC` | Base window handler | Fetches PowerPoint object model |
| `DocumentWindow` | Document management | Focus redirection, selection handling |
| `SlideBase/Slide/Master` | Slide representation | Accessibility info, slide properties |
| `Shape` | Generic shape handling | Role detection, positioning |
| `Table/TableCell` | Table support | Table structure navigation |
| `TextFrame` | Editable text | Custom TextInfo for navigation |
| `SlideShowWindow` | Slideshow mode | Notes support, caret navigation |

### Supported Features

**Working Well:**
- Slide navigation (Page Up/Down)
- Text editing in text boxes
- Shape name/type announcement
- Speaker notes in slideshow mode (with limitations)
- MathType equation reading
- Braille display routing (NVDA 2024.4+)

**Partially Working:**
- Chart reading (requires entering interaction mode)
- Table navigation (basic support)
- Grouped object handling (inconsistent)

**Not Working / Limited:**
- Outline view reading (Issue #12719)
- Grouped shapes with alt text (Issue #16161)
- Decorative element filtering (announced anyway)
- Comment navigation (no dedicated shortcuts)

### Event Handling

The app module uses COM event sink (`ppEApplicationSink`) to respond to:
- Slide changes
- Selection modifications
- Presentation opening/closing

---

## 4. Related Office Add-ons Analysis

### Office Desk (josephsl/officeDesk)

**Status:** End of Life (September 1, 2025)

**Features:**
- Backstage view search count announcement
- Prevents repetitive formatting announcements (bold/italic toggle)
- Labels edit fields in Envelopes dialog

**PowerPoint Support:** Originally planned but never implemented. The repository structure includes placeholder for PowerPoint app module, but no PowerPoint-specific code was developed.

**Relevance:** Architecture patterns could inform our plugin design (app module structure, settings organization).

**Repository:** https://github.com/josephsl/officeDesk

---

### wordAccessEnhancement (paulber19)

**Status:** Actively maintained (v3.7.2)

**Features:**
- Object navigation dialog (comments, revisions, bookmarks, fields, endnotes, footnotes, spelling/grammar errors)
- Cursor position announcement (line, column, page)
- Comment insertion shortcut
- Revision/footnote reading at cursor
- Sentence-by-sentence navigation
- Table navigation enhancements
- Browse mode command keys
- Spelling checker accessibility

**PowerPoint Applicability:** HIGH

The architecture and patterns in this add-on are directly applicable to PowerPoint:

| Word Feature | PowerPoint Equivalent |
|--------------|----------------------|
| Comment navigation | Comment navigation |
| Object listing dialog | Shape/image listing |
| Position announcement | Slide/shape position |
| Table navigation | Table cell navigation |
| Sentence navigation | Text frame navigation |

**Key Code Patterns to Reuse:**
- Object list dialog structure
- Navigation script patterns
- Position announcement formatting
- Browse mode key bindings

**Repository:** https://github.com/paulber19/wordAccessEnhancementNVDAAddon

---

## 5. Known PowerPoint Accessibility Issues

### Critical Issues

#### Issue #16161: Objects and Groups Not Perceived in Slideshow
**Impact:** High - Content inaccessible during presentations

**Symptoms:**
1. Grouped text boxes with shapes - text read but group alt text ignored
2. Grouped star and arrow shapes - completely ignored
3. Individual shapes with alt text - some ignored
4. SmartArt - alt text read but internal text ignored

**Root Cause:** Inconsistent handling of grouped objects and decorative markers in PowerPoint's accessibility API.

**JAWS Comparison:** JAWS handles these elements correctly, suggesting the information is available but NVDA isn't retrieving it.

---

#### Issue #12719: Outline View Not Reading
**Impact:** Medium - Alternative editing mode unusable

**Symptoms:** NVDA reads only "linefeed" and "blank" instead of actual slide text.

**Root Cause:** Outline view uses UI Automation which NVDA isn't capturing properly for this context.

**Status:** Open since 2021, no fix available.

---

### Moderate Issues

#### Comment Navigation Challenges
**Current State:** No dedicated comment navigation shortcuts in NVDA PowerPoint support.

**User Workaround:** Tab through slide elements, use PowerPoint's native comment pane.

**Opportunity:** Dedicated Ctrl+Alt+PageUp/PageDown shortcuts (as planned in our plugin).

---

#### Speaker Notes in Slideshow
**Current State:** Notes available but require manual navigation.

**JAWS Advantage:** JAWS has Ctrl+Shift+N to display notes in virtual viewer during slideshow.

**NVDA Limitation:** Users must use presenter view or external device for notes access.

---

### Minor Issues

- Text position reporting can be inaccurate (Issue #9941)
- Decorative elements still read aloud during manual navigation
- Table cell navigation lacks header awareness

---

## 6. Feature Comparison: NVDA vs JAWS for PowerPoint

| Feature | NVDA | JAWS | Notes |
|---------|------|------|-------|
| **Basic Navigation** | | | |
| Slide navigation | Yes | Yes | Both use Page Up/Down |
| Shape navigation | Tab | Tab | Similar implementation |
| Text editing | Yes | Yes | Both functional |
| **Advanced Features** | | | |
| Speaker notes (slideshow) | Limited | Ctrl+Shift+N | JAWS has virtual viewer |
| Comment navigation | No shortcuts | Yes | NVDA lacks dedicated commands |
| Chart interaction | Enter to interact | Similar | Both require interaction mode |
| MathType equations | Yes | Yes | Both support with MathType installed |
| **Object Handling** | | | |
| Grouped objects | Inconsistent | Better | JAWS reads group alt text |
| Decorative markers | Reads anyway | Respects | NVDA limitation |
| SmartArt text | Partial | Better | NVDA misses internal text |
| **Customization** | | | |
| Custom scripts | Python add-ons | JSL scripts | Different approaches |
| Cost | Free | $90-$1475/year | Significant cost difference |

### Key Differentiators

**NVDA Advantages:**
- Free and open-source
- Python-based customization
- Active community development
- Cross-platform (via Wine on Linux)

**JAWS Advantages:**
- More mature PowerPoint support
- Better speaker notes access
- More consistent object handling
- Enterprise support available

---

## 7. Gap Analysis

### Unmet User Needs

Based on research, the following needs are NOT addressed by existing NVDA plugins:

| Gap | User Impact | Priority | Our Plugin Addresses? |
|-----|-------------|----------|----------------------|
| Comment navigation shortcuts | Cannot efficiently review feedback | High | Yes (Ctrl+Alt+PgUp/PgDn) |
| Image navigation shortcuts | Manual search through shapes | High | Yes (Ctrl+Alt+Arrow) |
| AI image descriptions | No alt text = no access | Critical | Yes (Ollama/LLaVA) |
| Grouped object handling | Content missed in slideshows | High | Partial (via COM API) |
| Speaker notes quick access | Notes inaccessible during present | Medium | Future enhancement |
| Outline view support | Alternative editing blocked | Medium | Requires UIA work |
| Shape/object listing dialog | No overview of slide content | Medium | Future (Ctrl+Alt+C context) |
| Table header awareness | Lost in large tables | Medium | Future enhancement |

### Features Unique to Our Plugin

No existing NVDA add-on provides:

1. **AI-Powered Image Descriptions** - Local vision model integration for images without alt text
2. **Comment Navigation Shortcuts** - Dedicated keys for reviewing feedback
3. **Image-Specific Navigation** - Direct jumping between images
4. **Context Awareness** - Intelligent announcement of current position and content type
5. **Description Caching** - Performance optimization for repeated access

---

## 8. Recommendations

### Build vs Extend Analysis

| Option | Pros | Cons | Recommendation |
|--------|------|------|----------------|
| **Build New Plugin** | Full control, custom features, AI integration | More work, no existing user base | **RECOMMENDED** |
| **Extend Office Desk** | Existing structure | End of life, Word-focused | Not recommended |
| **Contribute to NVDA Core** | Benefits all users | Slower process, no AI features | Future consideration |

### Strategic Recommendations

#### Recommendation 1: Proceed with Custom Plugin Development
**Rationale:** No existing solution addresses the identified gaps. The wordAccessEnhancement patterns can inform design without requiring extension of that codebase.

#### Recommendation 2: Use COM API (Not UIA)
**Rationale:** NVDA's core PowerPoint support uses COM for good reason - it provides reliable access to PowerPoint's object model. UIA support in PowerPoint has documented issues.

#### Recommendation 3: Plan for NVDA Core Contribution
**Rationale:** Non-AI features (comment navigation, shape navigation) could eventually be contributed to NVDA core, benefiting all users. Keep code modular for this possibility.

#### Recommendation 4: Maintain JAWS Feature Parity Awareness
**Rationale:** Users comparing screen readers will expect feature parity. Document where NVDA+Plugin matches or exceeds JAWS capabilities.

#### Recommendation 5: Consider Collaboration with NV Access
**Rationale:** For features like better grouped object handling, collaboration with NV Access may be more effective than working around core limitations.

### Implementation Priority Matrix

| Feature | User Value | Implementation Effort | Priority |
|---------|------------|----------------------|----------|
| Image navigation | High | Low | P0 |
| AI image descriptions | Critical | Medium | P0 |
| Comment navigation | High | Low | P0 |
| Context awareness | Medium | Low | P1 |
| Description caching | Medium | Medium | P1 |
| Settings panel | Medium | Medium | P2 |
| Florence-2 backend | Medium | High | P2 |
| Shape listing dialog | Medium | Medium | P3 |
| Table enhancements | Low | High | P3 |

---

## 9. Contacts and Resources

### Key Developers/Contacts

| Name | Role | Contact | Notes |
|------|------|---------|-------|
| Joseph Lee (josephsl) | Office Desk author | GitHub: josephsl | Extensive Office knowledge |
| paulber19 | wordAccessEnhancement author | GitHub: paulber19 | Pattern reference |
| NV Access Team | NVDA core | GitHub: nvaccess | For core integration |
| Cyrille Bougot | Outlook Extended | GitHub: CyrilleB79 | Office add-on patterns |

### Community Resources

- **NVDA Users Mailing List:** nvda@nvda.groups.io
- **NVDA Add-ons Mailing List:** nvda-addons@groups.io
- **NVDA GitHub Issues:** https://github.com/nvaccess/nvda/issues
- **NV Access Community:** https://community.nvaccess.org/

### Training Resources

- **Microsoft PowerPoint with NVDA (eBook):** https://www.nvaccess.org/product/microsoft-powerpoint-with-nvda-ebook/
- **NVDA Developer Guide:** https://download.nvaccess.org/documentation/developerGuide.html
- **NVDA Add-on Development Guide:** https://github.com/nvdaaddons/DevGuide/wiki

---

## 10. References

### NVDA Add-on Stores
- NVDA Add-on Store: https://nvda.store/
- NVDA Add-ons Directory: https://nvda-addons.org/
- Legacy NVDA Add-ons: https://addons.nvda-project.org/

### GitHub Repositories
- NVDA Core: https://github.com/nvaccess/nvda
- Office Desk: https://github.com/josephsl/officeDesk
- wordAccessEnhancement: https://github.com/paulber19/wordAccessEnhancementNVDAAddon
- NVDA Add-on Template: https://github.com/accessolutions/nvda-addon-template

### Microsoft Documentation
- Screen Reader Support for PowerPoint: https://support.microsoft.com/en-us/office/screen-reader-support-for-powerpoint-9d2b646d-0b79-4135-a570-b8c7ad33ac2f
- Using Screen Reader to Navigate PowerPoint: https://support.microsoft.com/en-us/office/use-a-screen-reader-to-explore-and-navigate-powerpoint-a11115c7-6038-44a9-b355-8e133a4e9594
- Speaker Notes with Screen Reader: https://support.microsoft.com/en-us/office/use-a-screen-reader-to-read-or-add-speaker-notes-and-comments-in-powerpoint-0f40925d-8d78-4357-945b-ad7dd7bd7f60

### Historical Context
- Blind Orgs Make PowerPoint Support Reality: https://www.nvaccess.org/post/blind-orgs-make-powerpoint-support-a-reality-in-nvda/
- NFB: Creating PowerPoint with Screen Reader: https://nfb.org/resources/publications-and-media/access-podcast/creating-powerpoint-presentations-screen-reader

### Comparison Resources
- JAWS vs NVDA Comparison: https://blog.equally.ai/disability-guide/jaws-vs-nvda/
- UXPin Screen Reader Testing: https://www.uxpin.com/studio/blog/nvda-vs-jaws-screen-reader-testing-comparison/

---

## Document Information

| Field | Value |
|-------|-------|
| Document Title | NVDA PowerPoint Community Add-ons Research |
| Version | 1.0 |
| Date | 2025-12-04 |
| Author | Research Specialist Agent |
| Project | A11Y PowerPoint NVDA Plug-In |
| Status | Complete |

---

## Appendix A: Code Examples from wordAccessEnhancement

### Pattern: Object Navigation Dialog

```python
# Example pattern from wordAccessEnhancement for listing objects
# Applicable to PowerPoint shapes/comments listing

class ObjectListDialog(wx.Dialog):
    def __init__(self, parent, objects, title):
        super().__init__(parent, title=title)
        self.objects = objects
        self.listBox = wx.ListBox(self)
        for obj in objects:
            self.listBox.Append(obj.name)
        self.listBox.Bind(wx.EVT_LISTBOX_DCLICK, self.onSelect)

    def onSelect(self, event):
        index = self.listBox.GetSelection()
        if index != wx.NOT_FOUND:
            self.objects[index].navigate_to()
            self.Close()
```

### Pattern: Navigation Script with Position Announcement

```python
# Example pattern for announcing position during navigation
@script(
    description=_("Move to next comment"),
    gesture="kb:control+alt+pageDown"
)
def script_nextComment(self, gesture):
    comments = self.getComments()
    if not comments:
        ui.message(_("No comments on this slide"))
        return

    self.commentIndex = (self.commentIndex + 1) % len(comments)
    comment = comments[self.commentIndex]

    # Announce with position
    message = _("Comment {index} of {total}: {author} says: {text}").format(
        index=self.commentIndex + 1,
        total=len(comments),
        author=comment.author,
        text=comment.text
    )
    ui.message(message)
```

---

## Appendix B: NVDA PowerPoint App Module Structure

### Current Core Structure (powerpnt.py)

```
appModules/powerpnt.py
├── Classes
│   ├── PaneClassDC (base window handler)
│   ├── DocumentWindow (document management)
│   ├── OutlinePane
│   ├── SlideBase
│   │   ├── Slide
│   │   └── Master
│   ├── Shape
│   ├── Table
│   ├── TableCell
│   ├── TextFrame
│   ├── SlideShowWindow
│   └── SlideShowTreeInterceptor
├── Event Handlers
│   ├── ppEApplicationSink (COM events)
│   └── Selection change handlers
└── Gesture Bindings
    ├── Arrow key navigation
    ├── Tab navigation
    └── Page navigation
```

### Proposed Plugin Extension Structure

```
globalPlugins/powerpoint_ai/
├── __init__.py (package init)
├── navigation.py
│   ├── ImageNavigator
│   │   ├── next_image()
│   │   ├── previous_image()
│   │   └── get_all_images()
│   └── CommentNavigator
│       ├── next_comment()
│       ├── previous_comment()
│       └── get_all_comments()
├── image_analyzer.py
│   ├── OllamaAnalyzer
│   │   ├── describe_image()
│   │   └── check_status()
│   └── DescriptionCache
└── settings.py
    ├── SettingsPanel
    └── Configuration
```

---

*End of Research Document*
