# NVDA PowerPoint Comment Navigation Plugin - MVP Implementation Plan

## Project Overview

This plan outlines the implementation of an NVDA plugin focused on **enhanced comment navigation** for Microsoft PowerPoint 365. The AI-powered image features are deferred to a later phase.

**Target:** PowerPoint 365 with Modern Comments only (no legacy support needed)

**Repository:** This plugin is part of the `NVDA-Plugins` multi-plugin repository. See [REPO_STRUCTURE.md](REPO_STRUCTURE.md) for full details.

**Plugin Location:** `NVDA-Plugins/powerpoint-comments/`

---

## MVP Scope Summary

### HIGH PRIORITY (This Plan)
1. **View Management** - Detect and auto-switch to Normal view
2. **Slide Change Detection** - Automatic comment status announcement on slide change
3. **Focus First Comment** - Open Comments pane and focus first comment
4. **Slide Navigation from Comments** - Move to next/prev slide while in Comments pane
5. **@mention Detection** - Find comments mentioning current user
6. PowerPoint 365 Modern Comments only

### Core User Experience
```
User navigates to slide (PageUp/Down, click, etc.)
    ↓
Plugin detects slide change
    ↓
Plugin ensures Normal view (auto-switches if needed)
    ↓
Plugin checks for comments on current slide
    ↓
NO COMMENTS → Announce "No comments"
    ↓
HAS COMMENTS → Open Comments pane (if closed)
             → Announce "Has N comments"
             → Move focus to first comment
    ↓
User can then:
  - Navigate comments with Arrow keys (native behavior)
  - Change slides with shortcut (while in Comments pane)
  - Use native keys (Tab, Arrow) within Comments pane
```

### POST-MVP
- Slide comment summary (Ctrl+Alt+S)
- Comment-to-comment navigation (Ctrl+Alt+PageUp/Down) - if arrow keys prove insufficient

### BACKLOGGED
- Jump to unresolved comments (requires OOXML file parsing - complex file locking issues)

### LOW PRIORITY (Deferred)
- AI-powered image descriptions
- Image navigation
- Florence-2 model optimization

---

## Test Assets

### Available Test Resources
Located in `test_resources/`:

| Asset | Purpose |
|-------|---------|
| `Guide_Dogs_Test_Deck.pptx` | Main test presentation |
| `create_test_presentation.py` | Script to regenerate test deck |

### Test Deck Contents
- **9 slides** about guide dogs
- **1 TABLE** (Slide 3) - Country statistics
- **1 Text-based chart** (Slide 4) - Breed distribution
- **12 comments** across slides
- **Slide 6 has NO comments** - Tests empty case
- **Multiple @mentions** - Sarah Johnson, John Smith, Maria Garcia, David Chen

### Comment Distribution for Testing
| Slide | Title | Comments | @Mentions |
|-------|-------|----------|-----------|
| 1 | Title | 1 | @John Smith |
| 2 | What Are Guide Dogs | 2 | @Sarah Johnson |
| 3 | TABLE - Statistics | 2 | @Maria Garcia |
| 4 | Chart - Breeds | 1 | @John Smith |
| 5 | Training Process | 1 | @David Chen |
| 6 | Benefits | **0** | (none) |
| 7 | Etiquette | 3 | @John Smith, @Maria Garcia, @Sarah Johnson |
| 8 | Resources | 1 | (none) |
| 9 | Thank You | 1 | @John Smith, @Maria Garcia, @David Chen |

---

## Technical Approach Summary

Based on research findings:

| Component | Approach | Rationale |
|-----------|----------|-----------|
| COM Library | **comtypes** (not pywin32) | NVDA uses comtypes internally; pywin32 has DLL issues |
| View Detection | `ActiveWindow.ViewType` | Returns ppViewNormal (9) for Normal view |
| View Switching | `ActiveWindow.ViewType = 9` | Programmatically set Normal view |
| Slide Change | NVDA event tracking | Hook into gainFocus/slideChanged events |
| Comment Data | COM Automation | Access Slide.Comments collection |
| Focus Management | UIA SetFocus | Comments pane is UIA-enabled |
| @Mentions | Regex text parsing | No structured mention data in COM API |

---

## Phase 1: Foundation - App Module + View Management

**Goal:** Create working NVDA app module that connects to PowerPoint and manages view state

**Priority:** HIGHEST - Foundation for everything else

### 1.0 Logging Strategy (Critical for Debugging)

**Why:** Events may not fire as expected. Logging lets us verify event handling without visual feedback.

```python
import logging
log = logging.getLogger(__name__)

# Usage throughout module:
log.debug("Method called with args")   # Detailed tracing
log.info("Important event occurred")   # Key milestones
log.error(f"Operation failed: {e}")    # Errors
```

**View logs:** NVDA menu > Tools > View Log (or NVDA+F1)

**Key log points for Phase 1:**
- `__init__`: "PowerPoint Comments addon initialized"
- `event_appModule_gainFocus`: "event_appModule_gainFocus fired"
- `_connect_to_powerpoint`: "Connected to PowerPoint COM" or "Failed to connect: {error}"
- `_get_current_view`: "View type detected: {type}"
- `_ensure_normal_view`: "Switching view from {old} to Normal" or "Already in Normal view"

### 1.1 App Module Skeleton

```python
# appModules/powerpnt.py
# First, inherit all built-in PowerPoint support
from nvdaBuiltin.appModules.powerpnt import *

import appModuleHandler
from comtypes.client import GetActiveObject
import ui
import logging

log = logging.getLogger(__name__)

class AppModule(appModuleHandler.AppModule):
    """Enhanced PowerPoint with comment navigation."""

    # View type constants
    PP_VIEW_NORMAL = 9
    PP_VIEW_SLIDE_SORTER = 5
    PP_VIEW_NOTES = 10
    PP_VIEW_OUTLINE = 6
    PP_VIEW_SLIDE_MASTER = 3
    PP_VIEW_READING = 50

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._ppt_app = None
        self._last_slide_index = -1
        log.debug("PowerPoint Comments addon initialized")

    def event_appModule_gainFocus(self):
        """Called when PowerPoint gains focus."""
        log.debug("event_appModule_gainFocus fired")
        self._connect_to_powerpoint()
        self._ensure_normal_view()

    def _connect_to_powerpoint(self):
        """Connect to running PowerPoint instance."""
        try:
            self._ppt_app = GetActiveObject("PowerPoint.Application")
            log.debug("Connected to PowerPoint COM")
            return True
        except Exception as e:
            log.error(f"Failed to connect to PowerPoint: {e}")
            self._ppt_app = None
            return False

    def _get_current_view(self):
        """Get current PowerPoint view type."""
        try:
            if self._ppt_app and self._ppt_app.ActiveWindow:
                view_type = self._ppt_app.ActiveWindow.ViewType
                log.debug(f"View type detected: {view_type}")
                return view_type
        except Exception as e:
            log.error(f"Failed to get view type: {e}")
        return None

    def _ensure_normal_view(self):
        """Switch to Normal view if not already there."""
        try:
            current_view = self._get_current_view()
            if current_view is not None and current_view != self.PP_VIEW_NORMAL:
                log.info(f"Switching view from {current_view} to Normal")
                self._ppt_app.ActiveWindow.ViewType = self.PP_VIEW_NORMAL
                ui.message("Switched to Normal view")
                return True
            else:
                log.debug("Already in Normal view")
        except Exception as e:
            log.error(f"Failed to switch view: {e}")
        return False
```

### 1.2 COM Connection Verification

```python
def _verify_connection(self):
    """Verify COM connection is alive."""
    try:
        # Simple test - access ActivePresentation
        _ = self._ppt_app.ActivePresentation.Name
        return True
    except Exception:
        # Reconnect
        return self._connect_to_powerpoint()
```

### 1.3 Phase 1 Test Checklist

**Test Setup:**
1. Open NVDA
2. Open PowerPoint with `Guide_Dogs_Test_Deck.pptx`
3. Load plugin in NVDA scratchpad

**Tests:**
- [ ] Plugin loads without errors in NVDA log (`NVDA+F1` to open log)
- [ ] Focus PowerPoint → no errors
- [ ] Switch to Slide Sorter view manually → focus PowerPoint → auto-switches to Normal
- [ ] Switch to Reading view manually → focus PowerPoint → auto-switches to Normal
- [ ] Check NVDA log for "Switched to Normal view" message

**Phase 1 Exit Criteria:**
- Plugin loads cleanly
- View detection works
- Auto-switch to Normal view works

---

## Phase 1.1: Package + Deploy Pipeline

**Goal:** Create build and deployment pipeline so addon can be installed on test systems via GitHub

**Priority:** HIGH - Required for testing on separate system

### 1.1.1 Create Addon Directory Structure

```
powerpoint-comments/
├── addon/
│   ├── manifest.ini
│   └── appModules/
│       └── powerpnt.py
├── buildVars.py
└── README.md
```

### 1.1.2 manifest.ini

**CRITICAL: Follow quoting rules exactly!**

```ini
name = powerPointComments
summary = "Accessible PowerPoint Comment Navigation"
description = """Navigate and read PowerPoint comments with keyboard shortcuts and automatic announcements."""
author = "Electro Jam Instruments <contact@electrojam.com>"
url = https://github.com/Electro-Jam-Instruments/NVDAPlugIns/tree/main/powerpoint-comments
version = 0.1.0
minimumNVDAVersion = 2023.1
lastTestedNVDAVersion = 2024.4
```

### 1.1.3 Build Script

Create `build-tools/build_addon.py`:

```python
#!/usr/bin/env python
"""Build NVDA addon package."""
import os
import zipfile
from pathlib import Path

def build_addon(addon_dir: Path, output_name: str, version: str):
    """Create .nvda-addon from addon directory."""
    output_path = addon_dir.parent / f"{output_name}-{version}.nvda-addon"

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for file_path in addon_dir.rglob('*'):
            if file_path.is_file() and '__pycache__' not in str(file_path):
                arcname = file_path.relative_to(addon_dir)
                zf.write(file_path, arcname)

    print(f"Built: {output_path}")
    return output_path

if __name__ == "__main__":
    import sys
    plugin_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("powerpoint-comments")
    addon_path = plugin_dir / "addon"
    build_addon(addon_path, "powerpoint-comments", "0.1.0")
```

### 1.1.4 GitHub Release Process (Automated)

**See [RELEASE.md](RELEASE.md) for full release management documentation.**

**Version Update (Manual - only when requested):**
```bash
python build-tools/bump_version.py powerpoint-comments 0.0.1
git add powerpoint-comments/addon/manifest.ini
git commit -m "Bump powerpoint-comments to v0.0.1"
git push origin main
```

**Create Beta Release:**
```bash
git tag powerpoint-comments-v0.0.1-beta
git push origin powerpoint-comments-v0.0.1-beta
```

**Create Stable Release:**
```bash
git tag powerpoint-comments-v0.0.1
git push origin powerpoint-comments-v0.0.1
```

**What Happens Automatically:**
1. GitHub Actions triggers on tag push
2. Validates tag version matches manifest.ini
3. Builds .nvda-addon package
4. Creates GitHub release (pre-release for beta, stable for release)
5. Uploads addon file

**Download URL (after release):**
```
https://github.com/Electro-Jam-Instruments/NVDAPlugIns/releases/download/powerpoint-comments-v0.0.1-beta/powerpoint-comments-0.0.1.nvda-addon
```

### 1.1.5 Phase 1.1 Test Checklist

**Build Verification:**
- [ ] Build script runs without errors
- [ ] .nvda-addon file created
- [ ] File size reasonable (should be small, < 50KB)

**GitHub Release:**
- [ ] Release created with correct tag
- [ ] .nvda-addon file uploaded
- [ ] Download link works

**Installation on Test System:**
- [ ] Download .nvda-addon from GitHub
- [ ] Double-click to install
- [ ] NVDA prompts for installation confirmation
- [ ] Restart NVDA
- [ ] Open PowerPoint
- [ ] Check NVDA log for "PowerPoint Comments addon initialized"
- [ ] Verify Phase 1 functionality works on test system

**Phase 1.1 Exit Criteria:**
- Build pipeline works
- Addon installable from GitHub release
- Phase 1 features work on separate test system

---

## Phase 2: Slide Change Detection + Comment Status

**Goal:** Detect when user changes slides and announce comment status

**Priority:** HIGH - Core user experience

### 2.1 Slide Change Detection

```python
class AppModule(appModuleHandler.AppModule):
    # ... (from Phase 1)

    def _get_current_slide_index(self):
        """Get current slide index (1-based)."""
        try:
            if self._ppt_app and self._ppt_app.ActiveWindow:
                return self._ppt_app.ActiveWindow.View.Slide.SlideIndex
        except Exception:
            pass
        return -1

    def _check_slide_changed(self):
        """Check if slide has changed since last check."""
        current = self._get_current_slide_index()
        if current != self._last_slide_index and current > 0:
            self._last_slide_index = current
            return True
        return False

    def event_gainFocus(self, obj, nextHandler):
        """Called when focus changes within PowerPoint."""
        # Check for slide change
        if self._check_slide_changed():
            self._on_slide_changed()
        nextHandler()
```

### 2.2 Comment Status Announcement

```python
def _get_comments_on_current_slide(self):
    """Get all comments on current slide."""
    try:
        slide = self._ppt_app.ActiveWindow.View.Slide
        comments = []
        for comment in slide.Comments:
            comments.append({
                'text': comment.Text,
                'author': comment.Author,
                'datetime': comment.DateTime
            })
        return comments
    except Exception:
        return []

def _on_slide_changed(self):
    """Handle slide change event."""
    # Ensure Normal view
    self._ensure_normal_view()

    # Get comments
    comments = self._get_comments_on_current_slide()

    if not comments:
        ui.message("No comments")
    else:
        count = len(comments)
        ui.message(f"Has {count} comment{'s' if count != 1 else ''}")

        # Open Comments pane and focus first comment
        self._open_comments_pane()
        self._focus_first_comment()
```

### 2.3 Open Comments Pane

```python
def _open_comments_pane(self):
    """Open the Comments task pane if not visible."""
    try:
        # Try multiple command names (varies by Office version)
        for cmd in ["ReviewShowComments", "ShowComments", "CommentsPane"]:
            try:
                self._ppt_app.CommandBars.ExecuteMso(cmd)
                return True
            except Exception:
                continue
    except Exception:
        pass
    return False
```

### 2.4 Phase 2 Test Checklist

**Test Setup:**
1. Open `Guide_Dogs_Test_Deck.pptx`
2. Load plugin
3. Go to Slide 1

**Tests:**
- [ ] Press PageDown → hear "Has 2 comments" (Slide 2)
- [ ] Press PageDown → hear "Has 2 comments" (Slide 3 - table)
- [ ] Continue to Slide 6 → hear "No comments"
- [ ] Continue to Slide 7 → hear "Has 3 comments"
- [ ] Comments pane opens automatically when slide has comments
- [ ] Navigate backwards with PageUp → announcements still work

**Edge Cases:**
- [ ] Go to first slide, PageUp → stays on slide, no crash
- [ ] Go to last slide, PageDown → stays on slide, no crash
- [ ] Click thumbnail directly → announcement works

**Phase 2 Exit Criteria:**
- Slide changes detected reliably
- Comment count announced correctly
- "No comments" announced for empty slides
- Comments pane opens automatically

---

## Phase 3: Focus First Comment on Slide Change

**Goal:** When landing on slide with comments, move focus to first comment

**Priority:** HIGH - Core navigation experience

### 3.1 UIA Integration for Focus

```python
from comtypes.client import CreateObject
import comtypes.gen.UIAutomationClient as UIA

class CommentFocusManager:
    """Manages focus to comments via UI Automation."""

    def __init__(self):
        self._automation = CreateObject(
            "{ff48dba4-60ef-4201-aa87-54103eef594e}",  # CUIAutomation
            interface=UIA.IUIAutomation
        )

    def focus_first_comment(self, hwnd):
        """Focus the first comment in the Comments pane."""
        try:
            # Get root element from window handle
            root = self._automation.ElementFromHandle(hwnd)

            # Find Comments pane
            pane = self._find_comments_pane(root)
            if not pane:
                return False

            # Find first comment item
            first_comment = self._find_first_comment_item(pane)
            if first_comment:
                first_comment.SetFocus()
                return True

        except Exception:
            pass
        return False

    def _find_comments_pane(self, root):
        """Find Comments pane element."""
        # Try by name
        name_cond = self._automation.CreatePropertyCondition(
            UIA.UIA_NamePropertyId, "Comments"
        )
        pane = root.FindFirst(UIA.TreeScope_Descendants, name_cond)

        if not pane:
            # Try by automation ID
            id_cond = self._automation.CreatePropertyCondition(
                UIA.UIA_AutomationIdPropertyId, "CommentsPane"
            )
            pane = root.FindFirst(UIA.TreeScope_Descendants, id_cond)

        return pane

    def _find_first_comment_item(self, pane):
        """Find first comment list item."""
        # Look for ListItem or TreeItem control types
        list_item_cond = self._automation.CreatePropertyCondition(
            UIA.UIA_ControlTypePropertyId, UIA.UIA_ListItemControlTypeId
        )
        items = pane.FindAll(UIA.TreeScope_Descendants, list_item_cond)

        if items and items.Length > 0:
            return items.GetElement(0)

        return None
```

### 3.2 Integration with Slide Change

```python
def _focus_first_comment(self):
    """Focus first comment via UIA."""
    import win32gui

    # Get PowerPoint window handle
    hwnd = win32gui.GetForegroundWindow()

    # Small delay for Comments pane to appear
    import time
    time.sleep(0.2)

    # Try to focus
    if not self._comment_focus_manager.focus_first_comment(hwnd):
        # Fallback: just announce the first comment
        comments = self._get_comments_on_current_slide()
        if comments:
            ui.message(f"First comment by {comments[0]['author']}")
            ui.message(comments[0]['text'])
```

### 3.3 Phase 3 Test Checklist

**Test Setup:**
1. Open `Guide_Dogs_Test_Deck.pptx`
2. Close Comments pane if open
3. Go to Slide 1

**Tests:**
- [ ] PageDown to Slide 2 → Comments pane opens → focus lands on first comment
- [ ] NVDA announces first comment content automatically
- [ ] Tab key moves to Reply button / other controls in comment
- [ ] Down Arrow moves to second comment
- [ ] Escape or click slide → focus returns to slide content

**Tests with Comments Pane Already Open:**
- [ ] Open Comments pane manually first
- [ ] Navigate slides → focus still lands on first comment

**Phase 3 Exit Criteria:**
- Focus lands on first comment automatically
- NVDA announces comment via UIA (not plugin-generated speech)
- User can use native keys after focus lands

---

## Phase 4: Comment-to-Comment Navigation

**Goal:** Navigate between comments using Ctrl+Alt+PageUp/Down

**Priority:** MEDIUM - Enhances usability after Phase 3

### 4.1 Comment Navigation Class

```python
class CommentNavigator:
    """Navigate between comments on current slide."""

    def __init__(self, app_module):
        self._app = app_module
        self._comment_index = 0
        self._comments_cache = []
        self._cache_slide_index = -1

    def _refresh_cache(self):
        """Refresh comments cache if slide changed."""
        current_slide = self._app._get_current_slide_index()
        if current_slide != self._cache_slide_index:
            self._comments_cache = self._app._get_comments_on_current_slide()
            self._cache_slide_index = current_slide
            self._comment_index = 0

    def next_comment(self):
        """Navigate to next comment."""
        self._refresh_cache()

        if not self._comments_cache:
            ui.message("No comments on this slide")
            tones.beep(200, 100)
            return

        if self._comment_index >= len(self._comments_cache) - 1:
            ui.message("Last comment")
            tones.beep(880, 50)
            return

        self._comment_index += 1
        self._focus_comment_at_index(self._comment_index)

    def previous_comment(self):
        """Navigate to previous comment."""
        self._refresh_cache()

        if not self._comments_cache:
            ui.message("No comments on this slide")
            tones.beep(200, 100)
            return

        if self._comment_index <= 0:
            ui.message("First comment")
            tones.beep(880, 50)
            return

        self._comment_index -= 1
        self._focus_comment_at_index(self._comment_index)

    def _focus_comment_at_index(self, index):
        """Focus comment at given index via UIA."""
        # Similar to Phase 3 but target specific index
        pass
```

### 4.2 Keyboard Shortcuts

```python
from scriptHandler import script
import tones

class AppModule(appModuleHandler.AppModule):
    # ... (previous code)

    @script(
        description="Next comment",
        gesture="kb:control+alt+pageDown",
        category="PowerPoint Comments"
    )
    def script_nextComment(self, gesture):
        """Navigate to next comment."""
        self._comment_navigator.next_comment()

    @script(
        description="Previous comment",
        gesture="kb:control+alt+pageUp",
        category="PowerPoint Comments"
    )
    def script_previousComment(self, gesture):
        """Navigate to previous comment."""
        self._comment_navigator.previous_comment()

    @script(
        description="First comment",
        gesture="kb:control+alt+home",
        category="PowerPoint Comments"
    )
    def script_firstComment(self, gesture):
        """Navigate to first comment."""
        self._comment_navigator.first_comment()

    @script(
        description="Last comment",
        gesture="kb:control+alt+end",
        category="PowerPoint Comments"
    )
    def script_lastComment(self, gesture):
        """Navigate to last comment."""
        self._comment_navigator.last_comment()
```

### 4.3 Phase 4 Test Checklist

**Test Setup:**
1. Open `Guide_Dogs_Test_Deck.pptx`
2. Go to Slide 7 (has 3 comments)

**Tests:**
- [ ] Ctrl+Alt+PageDown → moves to next comment
- [ ] Ctrl+Alt+PageDown again → moves to third comment
- [ ] Ctrl+Alt+PageDown again → "Last comment" + beep (no wrap)
- [ ] Ctrl+Alt+PageUp → moves back to second comment
- [ ] Ctrl+Alt+Home → jumps to first comment
- [ ] Ctrl+Alt+End → jumps to last comment

**Boundary Tests:**
- [ ] On Slide 6 (no comments) → Ctrl+Alt+PageDown → "No comments on this slide" + beep
- [ ] On first comment → Ctrl+Alt+PageUp → "First comment" + beep

**Phase 4 Exit Criteria:**
- All navigation shortcuts work
- Boundary beeps occur (no wrap-around)
- Position maintained across slide changes

---

## Phase 5: @Mention Detection

**Goal:** Find and navigate to comments mentioning the current user

**Priority:** MEDIUM - Value-add feature

### 5.1 User Identity Detection

```python
import ctypes
import os

class CurrentUserDetector:
    """Detect current user identity."""

    def __init__(self):
        self._cached_identity = None

    def get_identity(self):
        """Get current user identity with caching."""
        if self._cached_identity:
            return self._cached_identity

        identity = {
            'display_name': None,
            'first_name': None,
            'email': None,
            'username': os.environ.get('USERNAME')
        }

        # Windows display name
        try:
            display_name = self._get_windows_display_name()
            if display_name:
                identity['display_name'] = display_name
                parts = display_name.split()
                identity['first_name'] = parts[0] if parts else None
        except Exception:
            pass

        self._cached_identity = identity
        return identity

    def _get_windows_display_name(self):
        """Get display name from Windows."""
        GetUserNameEx = ctypes.windll.secur32.GetUserNameExW
        NameDisplay = 3

        size = ctypes.pointer(ctypes.c_ulong(0))
        GetUserNameEx(NameDisplay, None, size)

        if size.contents.value == 0:
            return None

        name_buffer = ctypes.create_unicode_buffer(size.contents.value)
        GetUserNameEx(NameDisplay, name_buffer, size)

        return name_buffer.value if name_buffer.value else None
```

### 5.2 Mention Parser

```python
import re
from difflib import SequenceMatcher

class MentionParser:
    """Parse and match @mentions."""

    MENTION_PATTERN = re.compile(
        r'@([A-Za-z\u00C0-\u024F][\w\u00C0-\u024F]*'
        r'(?:[-\'][\w\u00C0-\u024F]+)?'
        r'(?:\s+[A-Za-z\u00C0-\u024F][\w\u00C0-\u024F]*)*)',
        re.UNICODE
    )

    @classmethod
    def extract_mentions(cls, text):
        """Extract @mentions from text."""
        if not text:
            return []
        return cls.MENTION_PATTERN.findall(text)

    @classmethod
    def mentions_user(cls, text, identity, threshold=0.85):
        """Check if text mentions the given user."""
        mentions = cls.extract_mentions(text)

        for mention in mentions:
            mention_lower = mention.lower()

            # Exact match on display name
            if identity.get('display_name'):
                if mention_lower == identity['display_name'].lower():
                    return True

            # First name match
            if identity.get('first_name'):
                if mention_lower == identity['first_name'].lower():
                    return True

            # Fuzzy match
            for name in [identity.get('display_name'), identity.get('first_name')]:
                if name:
                    ratio = SequenceMatcher(None, mention_lower, name.lower()).ratio()
                    if ratio >= threshold:
                        return True

        return False
```

### 5.3 Find My Mentions Feature

```python
@script(
    description="Find next comment mentioning me",
    gesture="kb:control+alt+m",
    category="PowerPoint Comments"
)
def script_findMyMention(self, gesture):
    """Navigate to next comment mentioning current user."""
    identity = self._user_detector.get_identity()

    # Search all slides for mentions
    mentions = self._find_all_mentions_of_user(identity)

    if not mentions:
        ui.message("No comments mention you in this presentation")
        return

    # Find next mention after current position
    next_mention = self._find_next_mention(mentions)

    if next_mention:
        # Navigate to that slide and comment
        self._navigate_to_mention(next_mention)
    else:
        ui.message("No more mentions found")
```

### 5.4 Phase 5 Test Checklist

**Test Setup:**
1. Note your Windows display name
2. Modify test deck to include @mention of your name (or test with existing names)

**Tests:**
- [ ] Ctrl+Alt+M → finds first comment mentioning you
- [ ] Ctrl+Alt+M again → finds next mention
- [ ] If no mentions → "No comments mention you"
- [ ] @FirstName works (partial match)
- [ ] @FullName works (exact match)

**Edge Cases:**
- [ ] Multiple mentions in same comment
- [ ] Mention in reply (not parent comment)
- [ ] Case-insensitive matching

**Phase 5 Exit Criteria:**
- User identity detected correctly
- Mentions found across all slides
- Navigation to mentioned comments works

---

## Phase 6: Polish, Error Handling, and Packaging

**Goal:** Robust error handling and NVDA addon packaging

**Priority:** Required for release

### 6.1 Error Handling

```python
def _safe_com_call(self, func, *args, fallback=None):
    """Safely execute COM call with reconnection."""
    try:
        return func(*args)
    except Exception:
        # Try to reconnect
        if self._connect_to_powerpoint():
            try:
                return func(*args)
            except Exception:
                pass
        return fallback
```

### 6.2 Error Messages

| Scenario | Message |
|----------|---------|
| PowerPoint not running | "PowerPoint not connected" |
| No presentation open | "No presentation open" |
| COM disconnection | (silent reconnect attempt) |
| UIA focus fails | Fallback to announce-only |

### 6.3 NVDA Addon Structure

See [REPO_STRUCTURE.md](REPO_STRUCTURE.md) for full repository layout.

**Plugin directory structure:**
```
powerpoint-comments/
├── README.md                # Plugin documentation
├── CHANGELOG.md             # Version history
├── buildVars.py             # Build configuration
├── addon/
│   ├── manifest.ini
│   ├── appModules/
│   │   └── powerpnt.py
│   ├── globalPlugins/
│   │   └── powerpoint_comments/
│   │       ├── __init__.py
│   │       ├── comment_navigator.py
│   │       ├── mention_parser.py
│   │       ├── user_identity.py
│   │       └── uia_focus.py
│   ├── doc/
│   │   └── en/
│   │       └── readme.html
│   └── locale/              # Future translations
└── tests/
    └── test_comment_detection.py
```

**Built output:** `powerpoint-comments-1.0.0.nvda-addon`

### 6.4 manifest.ini

**CRITICAL: Quoting rules matter! No quotes for single words, double quotes for text with spaces, triple quotes for multi-line.**

```ini
name = powerPointComments
summary = "Accessible PowerPoint Comment Navigation"
description = """Automatically announces comment status when changing slides. Navigate comments with keyboard shortcuts. Find @mentions of yourself."""
author = "Electro Jam Instruments <contact@electrojam.com>"
version = 1.0.0
url = https://github.com/Electro-Jam-Instruments/NVDAPlugIns/tree/main/powerpoint-comments
minimumNVDAVersion = 2023.1
lastTestedNVDAVersion = 2024.4
```

See `.agent/experts/nvda-plugins/nvda-plugins.md` for full quoting rules.

### 6.5 Release Process

**Tagging:**
```bash
git tag powerpoint-comments-v1.0.0
git push origin powerpoint-comments-v1.0.0
```

**Download URL:**
```
https://github.com/Electro-Jam-Instruments/NVDAPlugIns/releases/download/powerpoint-comments-v1.0.0/powerpoint-comments-1.0.0.nvda-addon
```

See [REPO_STRUCTURE.md](REPO_STRUCTURE.md) for complete release workflow.

### 6.6 Phase 6 Test Checklist

**Error Handling:**
- [ ] Close PowerPoint while plugin running → no crash
- [ ] Open PowerPoint → plugin reconnects
- [ ] Close presentation → "No presentation open"
- [ ] Open presentation → works again

**Packaging:**
- [ ] Create .nvda-addon file
- [ ] Install via NVDA addon manager
- [ ] Plugin loads correctly
- [ ] All features work after install

**Phase 6 Exit Criteria:**
- All error cases handled gracefully
- Addon installs cleanly
- Documentation complete

---

## Keyboard Shortcuts Summary

| Shortcut | Action |
|----------|--------|
| (automatic) | Announce comment status on slide change |
| Arrow keys | Navigate between comments (native behavior) |
| Ctrl+Alt+PageDown | Next slide (from Comments pane) |
| Ctrl+Alt+PageUp | Previous slide (from Comments pane) |
| Ctrl+Alt+M | Find next comment mentioning me |

---

## Phase Priority Summary

| Phase | Description | Priority | Dependencies |
|-------|-------------|----------|--------------|
| 1 | Foundation + View Management | HIGHEST | None |
| 1.1 | Package + Deploy Pipeline | HIGH | Phase 1 |
| 2 | Slide Change + Comment Status | HIGH | Phase 1.1 |
| 3 | Focus First Comment | HIGH | Phase 2 |
| 3.1 | Slide Navigation from Comments | HIGH | Phase 3 |
| 4 | @Mention Detection | MEDIUM | Phase 3.1 |
| 5 | Polish + Packaging | Required | All |
| 6 | Comment Navigation (if needed) | LOW | Post-MVP if arrow keys insufficient |

---

## Technical Dependencies

### Required (Included with NVDA)
- **comtypes** - COM automation
- **ctypes** - Windows API access

### No Additional Installs Required
- UIA access through comtypes
- User identity through Windows SecurLib
- Regex through Python standard library

---

## Success Criteria

### MVP Success
- [ ] Auto-announces comment status on slide change
- [ ] Auto-switches to Normal view
- [ ] Focus lands on first comment
- [ ] Navigate comments with Arrow keys (native)
- [ ] Navigate slides from Comments pane (Ctrl+Alt+PageUp/Down)
- [ ] Find comments mentioning current user

### Full Release
- [ ] Packaged as .nvda-addon
- [ ] Documentation complete
- [ ] Error handling robust
- [ ] Tested with screen reader users

---

## Appendix: Research References

### Research Documents
1. `research/NVDA_PowerPoint_Native_Support_Analysis.md`
2. `research/PowerPoint_Comment_Focus_Navigation_Research.md`
3. `research/powerpoint_mention_detection_research.md`
4. `research/NVDA_UIA_Deep_Research.md`
5. `research/PowerPoint-COM-Automation-Research.md`

### Key Findings
- NVDA has **NO** native comment support - clear opportunity
- Use **comtypes** not pywin32
- Comments pane is **UIA-enabled** (NetUIHWNDElement)
- @mentions are **plain text** in Comment.Text - parse with regex
- `ActiveWindow.ViewType` for view detection (Normal = 9)

---

**Document Version:** 3.0
**Last Updated:** December 2025
**Status:** Planning Complete - Ready for Phase 1 Implementation
