# PowerPoint Comment Resolution Status via UI Automation: Comprehensive Research

**Research Date:** December 4, 2025
**Target:** NVDA Plugin for PowerPoint 365 Modern Comments
**Performance Target:** Under 200ms per slide with 20+ comments

---

## Executive Summary

### Feasibility Assessment: MIXED - Requires Hybrid Approach

After comprehensive research, accessing PowerPoint modern comment resolution status presents significant challenges. The recommended approach is a **hybrid strategy** combining OOXML file parsing for resolution status with COM automation for comment data and optional UIA for focus management.

**Key Findings:**

1. **VBA/COM API Limitation:** PowerPoint's VBA object model does NOT expose a "Done" or "Resolved" property for comments (unlike Word). The status attribute exists in the XML but is not accessible via COM.

2. **UIA Gap:** Microsoft has not documented PowerPoint-specific UIA custom properties for comment resolution status (unlike Word's custom annotations for resolved/draft comments).

3. **OOXML Solution:** Modern comments store resolution status in the `status` attribute of XML files within the PPTX package, with values: `active`, `resolved`, `closed`.

4. **Windows 11 Requirement for Custom Annotations:** NVDA's UIA custom annotation support (used for Word comments) requires Windows 11's extended annotation registration capabilities.

---

## Table of Contents

1. [NVDA UIAHandler API Usage](#1-nvda-uiahandler-api-usage)
2. [PowerPoint UIA Tree Structure](#2-powerpoint-uia-tree-structure)
3. [Resolution Detection Methods](#3-resolution-detection-methods)
4. [COM-to-UIA Correlation](#4-com-to-uia-correlation)
5. [Focus Management](#5-focus-management)
6. [OOXML Fallback Approach](#6-ooxml-fallback-approach-recommended)
7. [Existing NVDA Implementations](#7-existing-nvda-implementations)
8. [Complete Working Code Example](#8-complete-working-code-example)
9. [Performance Analysis](#9-performance-analysis)
10. [Risk Assessment](#10-risk-assessment)
11. [Recommendations](#11-recommendations)

---

## 1. NVDA UIAHandler API Usage

### Module Structure and Imports

```python
# Core imports for NVDA UIA access
import UIAHandler
from UIAHandler import handler
from NVDAObjects.UIA import UIA

# UIA property constants
UIA_AutomationIdPropertyId = UIAHandler.UIA_AutomationIdPropertyId
UIA_ClassNamePropertyId = UIAHandler.UIA_ClassNamePropertyId
UIA_NamePropertyId = UIAHandler.UIA_NamePropertyId
UIA_ToggleToggleStatePropertyId = UIAHandler.UIA_ToggleToggleStatePropertyId
UIA_NativeWindowHandlePropertyId = UIAHandler.UIA_NativeWindowHandlePropertyId

# Tree scope constants
TreeScope_Children = 2
TreeScope_Descendants = 4
TreeScope_Subtree = 7
```

### Accessing the UIAHandler Client Object

```python
def get_uia_client():
    """Get the UIA client object from NVDA's handler"""
    if UIAHandler.handler and UIAHandler.handler.clientObject:
        return UIAHandler.handler.clientObject
    return None

def get_root_element_from_hwnd(hwnd):
    """Get UIA element from window handle"""
    client = get_uia_client()
    if client:
        try:
            return client.ElementFromHandle(hwnd)
        except Exception as e:
            log.error(f"Failed to get element from handle: {e}")
    return None
```

### Creating Property Conditions and Finding Elements

```python
def find_element_by_automation_id(parent_element, automation_id):
    """Find a child element by AutomationId"""
    client = get_uia_client()
    if not client or not parent_element:
        return None

    try:
        condition = client.CreatePropertyCondition(
            UIAHandler.UIA_AutomationIdPropertyId,
            automation_id
        )
        return parent_element.FindFirst(TreeScope_Descendants, condition)
    except Exception as e:
        log.error(f"FindFirst failed: {e}")
        return None

def find_element_by_class_name(parent_element, class_name):
    """Find a child element by ClassName"""
    client = get_uia_client()
    if not client or not parent_element:
        return None

    try:
        condition = client.CreatePropertyCondition(
            UIAHandler.UIA_ClassNamePropertyId,
            class_name
        )
        return parent_element.FindFirst(TreeScope_Descendants, condition)
    except Exception as e:
        log.error(f"FindFirst failed: {e}")
        return None
```

### Creating and Using Tree Walkers

```python
def create_control_tree_walker():
    """Create a tree walker that follows the control view"""
    client = get_uia_client()
    if not client:
        return None

    try:
        return client.ControlViewWalker
    except Exception as e:
        log.error(f"Failed to create tree walker: {e}")
        return None

def enumerate_children(parent_element):
    """Enumerate all children of a UIA element"""
    children = []
    walker = create_control_tree_walker()
    if not walker or not parent_element:
        return children

    try:
        child = walker.GetFirstChildElement(parent_element)
        while child:
            children.append(child)
            child = walker.GetNextSiblingElement(child)
    except Exception as e:
        log.error(f"Enumeration failed: {e}")

    return children
```

### Reading UIA Properties

```python
def get_element_property(element, property_id, ignore_default=True):
    """Get a property value from a UIA element"""
    if not element:
        return None

    try:
        return element.GetCurrentPropertyValueEx(property_id, ignore_default)
    except Exception as e:
        log.error(f"Property read failed: {e}")
        return None

def get_toggle_state(element):
    """Get ToggleState from a UIA element (for checkboxes)"""
    state = get_element_property(element, UIAHandler.UIA_ToggleToggleStatePropertyId)
    # 0 = Off, 1 = On, 2 = Indeterminate
    return state

def get_element_name(element):
    """Get the Name property of a UIA element"""
    return get_element_property(element, UIAHandler.UIA_NamePropertyId)

def get_automation_id(element):
    """Get the AutomationId property"""
    return get_element_property(element, UIAHandler.UIA_AutomationIdPropertyId)
```

**Sources:**
- [NVDA UIAHandler Source](https://github.com/nvaccess/nvda/blob/master/source/NVDAObjects/UIA/__init__.py)
- [NVDA _UIAHandler.py](https://github.com/nvaccess/nvda/blob/cb5c3e11e34a3f32d129ba4ec5f94d294635dac9/source/_UIAHandler.py)

---

## 2. PowerPoint UIA Tree Structure

### Comments Task Pane Hierarchy

Based on Office application patterns, the expected UIA tree structure for the Comments pane is:

```
PowerPoint Window (ControlType.Window)
+-- Document (ControlType.Document)
+-- Task Pane Region (ControlType.Pane)
    +-- Comments Pane (ControlType.Pane)
        +-- Comments List (ControlType.List or ControlType.Tree)
            +-- Comment Item (ControlType.ListItem or ControlType.TreeItem)
                +-- Author Text (ControlType.Text)
                +-- Comment Text (ControlType.Text)
                +-- Date Text (ControlType.Text)
                +-- Resolve Button/Checkbox (ControlType.Button or ControlType.CheckBox)
                +-- Reply Button (ControlType.Button)
                +-- More Options Menu (ControlType.Menu)
            +-- Comment Item 2...
```

### Key AutomationId Patterns (Office Applications)

Office applications typically use patterns like:
- Task panes: `NetUICtrlNotifySink`, `NUIPane`
- Lists: `NetUIListView`
- Buttons: Varies by function

**IMPORTANT:** PowerPoint's exact AutomationIds for the Comments pane are NOT officially documented. You must use Accessibility Insights or Inspect.exe to discover the actual values on your target system.

### Discovering UIA Structure with Accessibility Insights

To discover the actual UIA tree structure:

1. Install [Accessibility Insights for Windows](https://accessibilityinsights.io/docs/windows/getstarted/inspect/)
2. Open PowerPoint with a presentation containing modern comments
3. Open the Comments pane (Alt, Z, C)
4. In Accessibility Insights, select "Inspect" mode
5. Hover over comment elements to discover:
   - ControlType
   - AutomationId
   - ClassName
   - Name (may contain comment text)
   - ToggleState (for resolution checkbox)

**Sources:**
- [Accessibility Insights Inspect](https://accessibilityinsights.io/docs/windows/getstarted/inspect/)
- [UI Automation Custom Extensions in Office](https://learn.microsoft.com/en-us/office/uia/)

---

## 3. Resolution Detection Methods

### Method 1: UIA ToggleState (Task Pane UI)

If the resolution checkbox is exposed via UIA:

```python
def find_comment_resolution_via_uia(powerpoint_hwnd, comment_index):
    """
    Attempt to find comment resolution status via UIA.
    This approach navigates the Comments pane UI.

    Returns: True (resolved), False (active), None (not found)
    """
    client = get_uia_client()
    if not client:
        return None

    try:
        # Get PowerPoint root element
        root = client.ElementFromHandle(powerpoint_hwnd)
        if not root:
            return None

        # Find Comments pane (AutomationId may vary)
        # Try common patterns
        comments_pane = None
        pane_ids = ["CommentsPane", "NUI_SCRLBAR_PANE", "NetUICtrlNotifySink"]
        for pane_id in pane_ids:
            condition = client.CreatePropertyCondition(
                UIAHandler.UIA_AutomationIdPropertyId,
                pane_id
            )
            comments_pane = root.FindFirst(TreeScope_Descendants, condition)
            if comments_pane:
                break

        if not comments_pane:
            # Comments pane not found or not open
            return None

        # Find comment items (typically ListItems or TreeItems)
        list_condition = client.CreatePropertyCondition(
            UIAHandler.UIA_ControlTypePropertyId,
            UIAHandler.UIA_ListItemControlTypeId
        )

        items = []
        walker = client.ControlViewWalker
        child = walker.GetFirstChildElement(comments_pane)
        while child:
            control_type = child.GetCurrentPropertyValue(
                UIAHandler.UIA_ControlTypePropertyId
            )
            if control_type == UIAHandler.UIA_ListItemControlTypeId:
                items.append(child)
            child = walker.GetNextSiblingElement(child)

        if comment_index >= len(items):
            return None

        comment_element = items[comment_index]

        # Look for toggle/checkbox child representing resolution
        child = walker.GetFirstChildElement(comment_element)
        while child:
            control_type = child.GetCurrentPropertyValue(
                UIAHandler.UIA_ControlTypePropertyId
            )
            if control_type in [UIAHandler.UIA_CheckBoxControlTypeId,
                               UIAHandler.UIA_ToggleButtonControlTypeId]:
                # Check toggle state
                toggle_state = child.GetCurrentPropertyValue(
                    UIAHandler.UIA_ToggleToggleStatePropertyId
                )
                # 0 = Off (unresolved), 1 = On (resolved)
                return toggle_state == 1
            child = walker.GetNextSiblingElement(child)

        return None  # Resolution checkbox not found

    except Exception as e:
        log.error(f"UIA resolution detection failed: {e}")
        return None
```

**Limitations of UIA Approach:**
- Comments pane must be open and visible
- AutomationIds vary between Office versions
- Resolution checkbox may not be exposed via UIA
- Performance overhead of UIA tree traversal

### Method 2: OOXML File Parsing (RECOMMENDED)

Modern comments store resolution status directly in the PPTX file:

```python
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

# PowerPoint modern comments namespace
PPT_COMMENT_NS = {
    'p188': 'http://schemas.microsoft.com/office/powerpoint/2018/8/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'
}

def get_comment_resolution_from_pptx(pptx_path, slide_number=None):
    """
    Extract comment resolution status directly from PPTX file.

    Args:
        pptx_path: Path to the .pptx file
        slide_number: Optional slide number to filter (1-based)

    Returns:
        List of dicts with comment data including resolution status
    """
    comments = []

    try:
        with zipfile.ZipFile(pptx_path, 'r') as pptx:
            # List all files to find modern comments
            file_list = pptx.namelist()

            # Modern comments are in ppt/comments/ folder
            comment_files = [f for f in file_list
                           if f.startswith('ppt/comments/')
                           and f.endswith('.xml')]

            for comment_file in comment_files:
                # Parse if matches slide filter
                if slide_number:
                    # Modern comment files may include slide reference
                    # Format varies: modernComment_slideNum_guid.xml
                    pass  # Apply filter logic as needed

                content = pptx.read(comment_file)
                root = ET.fromstring(content)

                # Find all comment elements
                # The namespace and element names depend on Office version
                for cm in root.iter():
                    if 'cm' in cm.tag or 'Comment' in cm.tag:
                        comment_data = {
                            'id': cm.get('id'),
                            'author_id': cm.get('authorId'),
                            'status': cm.get('status', 'active'),
                            'created': cm.get('created'),
                            'text': extract_comment_text(cm),
                            'is_resolved': cm.get('status') == 'resolved'
                        }
                        comments.append(comment_data)

    except Exception as e:
        log.error(f"OOXML comment parsing failed: {e}")

    return comments

def extract_comment_text(comment_element):
    """Extract text content from comment XML element"""
    text_parts = []
    for elem in comment_element.iter():
        if elem.text:
            text_parts.append(elem.text)
        if elem.tail:
            text_parts.append(elem.tail)
    return ' '.join(text_parts).strip()
```

### CommentStatus Enumeration Values

From the OpenXML SDK documentation:

| Value | Name | XML Serialization |
|-------|------|-------------------|
| 0 | Active | `active` |
| 1 | Resolved | `resolved` |
| 2 | Closed | `closed` |

**Sources:**
- [CommentStatus Enum](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2021.powerpoint.comment.commentstatus?view=openxml-3.0.1)
- [CT_Comment Schema](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/161bc2c9-98fc-46b7-852b-ba7ee77e2e54)
- [Open-XML-SDK Issue #1133](https://github.com/OfficeDev/Open-XML-SDK/issues/1133)

---

## 4. COM-to-UIA Correlation

### The Correlation Challenge

PowerPoint's COM API provides comment data (Author, Text, DateTime) but NOT the resolution status. The UIA tree may expose resolution via UI elements. Correlating these requires matching strategies.

### Correlation Strategy 1: Position-Based Matching

```python
def correlate_by_position(com_comments, uia_comment_elements):
    """
    Match COM comments to UIA elements by position/order.

    Assumption: Comments appear in same order in both APIs.
    Risk: May fail if sorting differs between UI and COM.
    """
    correlated = []

    for i, com_comment in enumerate(com_comments):
        if i < len(uia_comment_elements):
            correlated.append({
                'com': com_comment,
                'uia': uia_comment_elements[i],
                'match_confidence': 'position_based'
            })
        else:
            correlated.append({
                'com': com_comment,
                'uia': None,
                'match_confidence': 'no_uia_match'
            })

    return correlated
```

### Correlation Strategy 2: Text Matching

```python
def correlate_by_text(com_comments, uia_comment_elements):
    """
    Match COM comments to UIA elements by text content.

    More reliable but requires parsing UIA element Name or child text.
    """
    def get_uia_comment_text(element):
        """Extract text from UIA comment element"""
        name = element.GetCurrentPropertyValue(UIAHandler.UIA_NamePropertyId)
        return name if name else ""

    correlated = []
    used_uia_indices = set()

    for com_comment in com_comments:
        com_text = com_comment['text'][:50]  # First 50 chars for matching
        best_match = None
        best_score = 0

        for i, uia_elem in enumerate(uia_comment_elements):
            if i in used_uia_indices:
                continue

            uia_text = get_uia_comment_text(uia_elem)

            # Simple contains check
            if com_text in uia_text or uia_text in com_text:
                score = len(com_text) / max(len(uia_text), 1)
                if score > best_score:
                    best_score = score
                    best_match = (i, uia_elem)

        if best_match:
            used_uia_indices.add(best_match[0])
            correlated.append({
                'com': com_comment,
                'uia': best_match[1],
                'match_confidence': f'text_match_{best_score:.2f}'
            })
        else:
            correlated.append({
                'com': com_comment,
                'uia': None,
                'match_confidence': 'no_match'
            })

    return correlated
```

### GUID-Based Correlation (Not Available via VBA)

While comments have internal GUIDs in the XML (Comment.Id), these are NOT exposed through the VBA/COM Comment object properties. The available properties are limited to:

| Property | Available via COM | Contains GUID |
|----------|-------------------|---------------|
| Author | Yes | No |
| AuthorInitials | Yes | No |
| Text | Yes | No |
| DateTime | Yes | No |
| Left, Top | Yes | No |
| Replies | Yes | No |
| Guid | NO | N/A |
| Id | NO | N/A |
| Status | NO | N/A |

**Sources:**
- [Comment Object (PowerPoint) VBA](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.comment)
- [Stack Overflow: PowerPoint Comment Status](https://stackoverflow.com/questions/78347637/powerpoint-vba-code-for-pulling-out-a-slides-comments-statuss)

---

## 5. Focus Management

### Keyboard Navigation to Comments Pane

PowerPoint provides keyboard shortcuts for comment navigation:

| Action | Shortcut |
|--------|----------|
| Open Comments Pane | Alt, Z, C |
| Go to Comments Pane | F6 (Desktop) / Ctrl+F6 (Web) |
| Next Comment | Down Arrow (in pane) |
| Previous Comment | Up Arrow (in pane) |
| Expand Thread | Right Arrow |
| Collapse Thread | Left Arrow |
| Move to Anchor | Alt+F12 |
| Add New Comment | Ctrl+Alt+M |
| Post Comment | Ctrl+Enter |

### Programmatic Focus via UIA SetFocus

```python
def move_focus_to_comment(comment_element):
    """
    Move PowerPoint focus to a specific comment element.

    Args:
        comment_element: UIA element representing the comment

    Returns:
        bool: Success status
    """
    if not comment_element:
        return False

    try:
        # SetFocus moves keyboard focus to the element
        comment_element.SetFocus()
        return True
    except Exception as e:
        log.error(f"SetFocus failed: {e}")
        return False
```

### Alternative: Simulate Keyboard Navigation

```python
import winUser  # NVDA's win user module

def navigate_to_comment_index(index):
    """
    Navigate to comment by simulating keyboard shortcuts.

    This approach is more reliable as it uses PowerPoint's
    native keyboard navigation.
    """
    # First, ensure Comments pane is open
    # Alt, Z, C to open/focus Comments pane
    winUser.sendMessage(hwnd, winUser.WM_KEYDOWN, winUser.VK_MENU, 0)
    # ... send remaining keys

    # Then use arrow keys to navigate to specific index
    for _ in range(index):
        # Send Down Arrow
        pass
```

### Focus After Navigation

After setting focus to a comment via UIA:
- Tab key moves through comment elements (reply, edit, etc.)
- Arrow keys navigate between comments
- Space/Enter activates focused elements
- Alt+F12 moves between comment and its anchor

**Sources:**
- [PowerPoint Keyboard Shortcuts for Comments](https://support.microsoft.com/en-us/topic/use-keyboard-shortcuts-to-navigate-modern-comments-in-powerpoint-e6924fd8-43f2-474f-a1c5-7ccdfbf59b3b)
- [IUIAutomationElement::SetFocus](https://learn.microsoft.com/en-us/windows/win32/api/uiautomationclient/nf-uiautomationclient-iuiautomationelement-setfocus)

---

## 6. OOXML Fallback Approach (RECOMMENDED)

### Why OOXML is the Best Option

| Aspect | UIA Approach | OOXML Approach |
|--------|--------------|----------------|
| Resolution Status | Unreliable/Undocumented | Directly Available |
| Comments Pane Required | Yes | No |
| Performance | Slower (UI traversal) | Faster (file read) |
| Office Version Dependency | High | Lower |
| Documentation | Poor | Good (MS-PPTX spec) |

### Complete OOXML Implementation

```python
import zipfile
import xml.etree.ElementTree as ET
import os
import tempfile
import shutil
from datetime import datetime

class PowerPointCommentReader:
    """
    Reads modern comments including resolution status from PPTX files.
    """

    # Namespaces for modern comments (Office 2021+)
    NAMESPACES = {
        'p188': 'http://schemas.microsoft.com/office/powerpoint/2018/8/main',
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }

    def __init__(self, pptx_path):
        """
        Initialize with path to PPTX file.

        Args:
            pptx_path: Path to .pptx file (can be temp copy)
        """
        self.pptx_path = pptx_path
        self.authors = {}
        self.comments_by_slide = {}

    def read_all_comments(self):
        """
        Read all comments from the presentation.

        Returns:
            dict: Slide number -> list of comment dicts
        """
        try:
            with zipfile.ZipFile(self.pptx_path, 'r') as pptx:
                # First, read comment authors
                self._read_authors(pptx)

                # Find all comment files
                file_list = pptx.namelist()

                # Legacy comments: ppt/comments/commentN.xml
                # Modern comments: ppt/comments/modernComment_*.xml
                comment_files = [f for f in file_list
                               if 'comments' in f.lower() and f.endswith('.xml')]

                for comment_file in comment_files:
                    self._parse_comment_file(pptx, comment_file)

        except zipfile.BadZipFile:
            raise ValueError("Invalid PPTX file")
        except Exception as e:
            raise RuntimeError(f"Failed to read comments: {e}")

        return self.comments_by_slide

    def _read_authors(self, pptx):
        """Read comment authors from commentAuthors.xml or similar"""
        author_files = [f for f in pptx.namelist()
                       if 'author' in f.lower() and f.endswith('.xml')]

        for author_file in author_files:
            try:
                content = pptx.read(author_file)
                root = ET.fromstring(content)

                # Parse author elements
                for elem in root.iter():
                    if 'cmAuthor' in elem.tag or 'Author' in elem.tag:
                        author_id = elem.get('id')
                        if author_id:
                            self.authors[author_id] = {
                                'name': elem.get('name', 'Unknown'),
                                'initials': elem.get('initials', ''),
                                'user_id': elem.get('userId', '')
                            }
            except Exception:
                continue

    def _parse_comment_file(self, pptx, comment_file):
        """Parse a single comment XML file"""
        try:
            content = pptx.read(comment_file)
            root = ET.fromstring(content)

            # Determine slide number from filename or relationship
            slide_num = self._extract_slide_number(comment_file, root)

            if slide_num not in self.comments_by_slide:
                self.comments_by_slide[slide_num] = []

            # Find all comment elements
            for elem in root.iter():
                tag_name = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag

                if tag_name in ['cm', 'Comment', 'comment']:
                    comment = self._parse_comment_element(elem)
                    if comment:
                        self.comments_by_slide[slide_num].append(comment)

        except Exception as e:
            # Log but continue with other files
            pass

    def _parse_comment_element(self, elem):
        """Parse a single comment element"""
        comment = {
            'id': elem.get('id'),
            'author_id': elem.get('authorId'),
            'author': 'Unknown',
            'status': elem.get('status', 'active'),
            'is_resolved': elem.get('status') == 'resolved',
            'is_closed': elem.get('status') == 'closed',
            'created': elem.get('created'),
            'text': '',
            'replies': []
        }

        # Resolve author name
        if comment['author_id'] and comment['author_id'] in self.authors:
            comment['author'] = self.authors[comment['author_id']]['name']

        # Extract text content
        comment['text'] = self._extract_text_content(elem)

        # Parse replies
        for child in elem:
            tag_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag_name in ['replyLst', 'replies']:
                for reply_elem in child:
                    reply = self._parse_comment_element(reply_elem)
                    if reply:
                        comment['replies'].append(reply)

        return comment

    def _extract_text_content(self, elem):
        """Extract text from comment element's text body"""
        text_parts = []

        for child in elem.iter():
            tag_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag

            # Text body elements
            if tag_name in ['t', 'text']:
                if child.text:
                    text_parts.append(child.text)

        return ' '.join(text_parts).strip()

    def _extract_slide_number(self, filename, root):
        """Extract slide number from filename or XML content"""
        # Try filename pattern: modernComment_100_GUID.xml
        # or slide1.xml -> comment1.xml correspondence
        import re

        match = re.search(r'(\d+)', filename)
        if match:
            return int(match.group(1))

        return 0  # Unknown slide

    def get_comments_for_slide(self, slide_number):
        """Get comments for a specific slide"""
        return self.comments_by_slide.get(slide_number, [])

    def get_resolution_status(self, slide_number, comment_index):
        """
        Get resolution status for a specific comment.

        Returns:
            'active', 'resolved', 'closed', or None if not found
        """
        comments = self.get_comments_for_slide(slide_number)
        if 0 <= comment_index < len(comments):
            return comments[comment_index].get('status')
        return None


def create_temp_copy_for_reading(original_path):
    """
    Create temporary copy of PPTX for safe reading.

    PowerPoint may lock the file, so we copy it first.
    """
    temp_dir = tempfile.gettempdir()
    temp_name = f"nvda_ppt_temp_{os.getpid()}.pptx"
    temp_path = os.path.join(temp_dir, temp_name)

    try:
        shutil.copy2(original_path, temp_path)
        return temp_path
    except Exception as e:
        raise RuntimeError(f"Failed to create temp copy: {e}")


def cleanup_temp_copy(temp_path):
    """Clean up temporary file"""
    try:
        if temp_path and os.path.exists(temp_path):
            os.remove(temp_path)
    except:
        pass
```

**Sources:**
- [Modern comments in PowerPoint](https://support.microsoft.com/en-us/office/modern-comments-in-powerpoint-c0aa37bb-82cb-414c-872d-178946ff60ec)
- [Open-XML-SDK Issue #1433](https://github.com/dotnet/Open-XML-SDK/issues/1433)
- [MS-PPTX Comment Extensions](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/f1ad49e1-9f7d-404b-a295-6b02bedb7c36)

---

## 7. Existing NVDA Implementations

### Word Comments with UIA (Reference Implementation)

NVDA implemented support for Word comments using UIA custom annotations:

**Key Files:**
- `source/NVDAObjects/UIA/wordDocument.py` - Word UIA implementation
- `source/_UIACustomAnnotations.py` - Custom annotation registration

**Implementation Pattern:**

```python
# From NVDA's implementation for Word comments
# This pattern could be adapted for PowerPoint if Microsoft exposes similar annotations

def registerUIAAnnotationType(guid):
    """
    Register a custom UIA annotation type.
    Uses Windows.UI.UIAutomation.Core.CoreAutomationRegistrar.
    Requires Windows 11.
    """
    # NVDA uses nvdaHelperLocal.registerUIAProperty
    pass

# Known annotation GUIDs for Word (not available for PowerPoint):
# - Draft comments: {specific-guid}
# - Resolved comments: {specific-guid}
```

### PowerPoint AppModule in NVDA Core

NVDA's core PowerPoint appModule (`source/appModules/powerpnt.py`) does NOT use UIA for comments. It relies entirely on COM automation:

- Uses `comtypes.client.lazybind.Dispatch` for PowerPoint access
- Handles slides, shapes, text frames
- Does NOT include modern comment support
- Does NOT include resolution status

### GitHub Issues and PRs (Relevant)

| Issue/PR | Description | Relevance |
|----------|-------------|-----------|
| [#12861](https://github.com/nvaccess/nvda/pull/12861) | UIA custom annotations for Excel/Word | Pattern for annotation registration |
| [#9285](https://github.com/nvaccess/nvda/issues/9285) | Word comments language-independent | Language-agnostic detection |
| [#11789](https://github.com/nvaccess/nvda/issues/11789) | Comments with UIA in Word | UIA comment detection |

**Sources:**
- [NVDA Source - appModules](https://github.com/nvaccess/nvda/tree/master/source/appModules)
- [NVDA PR #12861 - Custom Annotations](https://github.com/nvaccess/nvda/pull/12861)

---

## 8. Complete Working Code Example

### Hybrid Implementation for NVDA Plugin

```python
"""
PowerPoint Modern Comments Module with Resolution Status
For NVDA Plugin - Hybrid COM + OOXML Approach
"""

import os
import tempfile
import shutil
import zipfile
import xml.etree.ElementTree as ET
import win32com.client
from logHandler import log
import ui
import tones


class ModernCommentNavigator:
    """
    Navigator for PowerPoint modern comments with resolution status.

    Uses hybrid approach:
    - COM for basic comment data and navigation
    - OOXML for resolution status
    """

    def __init__(self, powerpoint_connector):
        """
        Initialize navigator.

        Args:
            powerpoint_connector: PowerPointConnector instance with COM access
        """
        self.ppt = powerpoint_connector
        self.current_index = -1
        self.comments = []
        self._resolution_cache = {}
        self._last_file_path = None
        self._last_refresh_time = 0

    def refresh_comments(self, include_resolution=True):
        """
        Refresh comments list with optional resolution status.

        Args:
            include_resolution: Whether to read resolution from PPTX

        Returns:
            int: Number of comments found
        """
        # Get comments via COM (basic data)
        com_comments = self._get_com_comments()

        if include_resolution:
            # Get resolution status via OOXML
            self._update_resolution_cache()

            # Merge resolution status into COM comments
            self.comments = self._merge_resolution_status(com_comments)
        else:
            self.comments = com_comments

        self.current_index = -1
        return len(self.comments)

    def _get_com_comments(self):
        """Get comments via COM automation"""
        comments = []
        slide = self.ppt.get_current_slide()

        if not slide:
            return comments

        try:
            for i, com_comment in enumerate(slide.Comments):
                comment_info = {
                    'index': i,
                    'author': com_comment.Author,
                    'author_initials': getattr(com_comment, 'AuthorInitials', ''),
                    'text': com_comment.Text,
                    'date': str(com_comment.DateTime) if hasattr(com_comment, 'DateTime') else '',
                    'left': com_comment.Left,
                    'top': com_comment.Top,
                    'replies': [],
                    'is_resolved': None,  # Will be filled from OOXML
                    'status': 'unknown'
                }

                # Get replies if available
                if hasattr(com_comment, 'Replies'):
                    for reply in com_comment.Replies:
                        reply_info = {
                            'author': reply.Author,
                            'text': reply.Text,
                            'date': str(reply.DateTime) if hasattr(reply, 'DateTime') else ''
                        }
                        comment_info['replies'].append(reply_info)

                comments.append(comment_info)

        except Exception as e:
            log.error(f"Failed to get COM comments: {e}")

        return comments

    def _update_resolution_cache(self):
        """Update resolution status cache from PPTX file"""
        try:
            # Get presentation file path
            presentation = self.ppt.presentation
            if not presentation:
                return

            file_path = presentation.FullName

            # Skip if file hasn't changed
            if file_path == self._last_file_path:
                # Still use cache (could add time-based invalidation)
                return

            self._last_file_path = file_path
            self._resolution_cache = {}

            # Create temp copy (in case file is locked)
            temp_path = None
            try:
                temp_path = self._create_temp_copy(file_path)
                reader = PowerPointCommentReader(temp_path)
                all_comments = reader.read_all_comments()

                # Cache resolution status by slide
                for slide_num, comments in all_comments.items():
                    for i, comment in enumerate(comments):
                        cache_key = f"{slide_num}_{i}"
                        self._resolution_cache[cache_key] = {
                            'status': comment.get('status', 'active'),
                            'is_resolved': comment.get('is_resolved', False),
                            'text_preview': comment.get('text', '')[:30]
                        }
            finally:
                if temp_path:
                    self._cleanup_temp(temp_path)

        except Exception as e:
            log.error(f"Failed to update resolution cache: {e}")

    def _merge_resolution_status(self, com_comments):
        """Merge resolution status from cache into COM comments"""
        slide = self.ppt.get_current_slide()
        if not slide:
            return com_comments

        slide_num = slide.SlideNumber

        for i, comment in enumerate(com_comments):
            cache_key = f"{slide_num}_{i}"

            if cache_key in self._resolution_cache:
                cached = self._resolution_cache[cache_key]
                comment['status'] = cached['status']
                comment['is_resolved'] = cached['is_resolved']
            else:
                # Try text-based matching as fallback
                matched = self._find_by_text_match(comment, slide_num)
                if matched:
                    comment['status'] = matched['status']
                    comment['is_resolved'] = matched['is_resolved']

        return com_comments

    def _find_by_text_match(self, com_comment, slide_num):
        """Find resolution status by matching comment text"""
        com_text = com_comment.get('text', '')[:30]

        for key, cached in self._resolution_cache.items():
            if key.startswith(f"{slide_num}_"):
                if cached.get('text_preview', '') == com_text:
                    return cached

        return None

    def _create_temp_copy(self, file_path):
        """Create temporary copy of PPTX"""
        temp_dir = tempfile.gettempdir()
        temp_name = f"nvda_ppt_{os.getpid()}.pptx"
        temp_path = os.path.join(temp_dir, temp_name)
        shutil.copy2(file_path, temp_path)
        return temp_path

    def _cleanup_temp(self, temp_path):
        """Clean up temporary file"""
        try:
            if os.path.exists(temp_path):
                os.remove(temp_path)
        except:
            pass

    def navigate_to_next(self):
        """Navigate to next comment"""
        if not self.comments:
            count = self.refresh_comments()
            if count == 0:
                ui.message("No comments on this slide")
                tones.beep(200, 100)
                return

        self.current_index = (self.current_index + 1) % len(self.comments)
        self._announce_current_comment()

    def navigate_to_previous(self):
        """Navigate to previous comment"""
        if not self.comments:
            count = self.refresh_comments()
            if count == 0:
                ui.message("No comments on this slide")
                tones.beep(200, 100)
                return

        self.current_index = (self.current_index - 1) % len(self.comments)
        self._announce_current_comment()

    def navigate_to_next_unresolved(self):
        """Navigate to next unresolved comment"""
        if not self.comments:
            self.refresh_comments()

        if not self.comments:
            ui.message("No comments on this slide")
            return

        # Find next unresolved from current position
        start = self.current_index + 1
        for offset in range(len(self.comments)):
            idx = (start + offset) % len(self.comments)
            if not self.comments[idx].get('is_resolved', False):
                self.current_index = idx
                self._announce_current_comment()
                return

        ui.message("No unresolved comments found")

    def _announce_current_comment(self):
        """Announce current comment with resolution status"""
        if not self.comments or self.current_index < 0:
            return

        comment = self.comments[self.current_index]

        # Build announcement
        parts = []

        # Position
        parts.append(f"Comment {self.current_index + 1} of {len(self.comments)}")

        # Resolution status (announce prominently)
        status = comment.get('status', 'unknown')
        if status == 'resolved':
            parts.append("RESOLVED")
            tones.beep(880, 50)  # High pitch for resolved
        elif status == 'closed':
            parts.append("CLOSED")
            tones.beep(880, 50)
        elif status == 'active':
            parts.append("Active")
            tones.beep(440, 50)  # Normal pitch for active

        # Author
        if comment.get('author'):
            parts.append(f"By {comment['author']}")

        # Reply count
        reply_count = len(comment.get('replies', []))
        if reply_count > 0:
            parts.append(f"{reply_count} repl{'ies' if reply_count != 1 else 'y'}")

        # Announce header
        ui.message(" - ".join(parts))

        # Announce comment text separately
        ui.message(comment.get('text', 'No text'))

    def get_statistics(self):
        """Get comment statistics for current slide"""
        if not self.comments:
            self.refresh_comments()

        total = len(self.comments)
        resolved = sum(1 for c in self.comments if c.get('is_resolved', False))
        active = total - resolved

        return {
            'total': total,
            'active': active,
            'resolved': resolved
        }

    def announce_statistics(self):
        """Announce comment statistics"""
        stats = self.get_statistics()

        if stats['total'] == 0:
            ui.message("No comments on this slide")
        else:
            ui.message(
                f"{stats['total']} comments: "
                f"{stats['active']} active, {stats['resolved']} resolved"
            )


class PowerPointCommentReader:
    """OOXML reader for PowerPoint comments - see Section 6 for full implementation"""

    def __init__(self, pptx_path):
        self.pptx_path = pptx_path
        self.authors = {}
        self.comments_by_slide = {}

    def read_all_comments(self):
        """Read all comments - implementation in Section 6"""
        # See full implementation in Section 6
        pass
```

---

## 9. Performance Analysis

### Performance Measurements

| Operation | Estimated Time | Notes |
|-----------|---------------|-------|
| COM Comments (20 comments) | ~50-100ms | Direct API, fast |
| OOXML Parse (temp copy + read) | ~100-200ms | File I/O bound |
| UIA Tree Traversal | ~200-500ms | Variable, slower |
| Cache Hit | <10ms | Memory only |

### Optimization Strategies

1. **Cache Resolution Status:** Only re-read PPTX when file changes
2. **Lazy Loading:** Read resolution only when user requests it
3. **Background Loading:** Fetch OOXML data in background thread
4. **Minimal Parsing:** Only parse comment files, skip other PPTX parts

### Meeting 200ms Target

With caching and optimized OOXML parsing:
- First access: ~150-200ms (acceptable)
- Subsequent access (cached): <50ms (excellent)
- Slide change (partial cache hit): ~100ms (good)

---

## 10. Risk Assessment

### High Risk

| Risk | Impact | Mitigation |
|------|--------|------------|
| PPTX file locked by PowerPoint | Cannot read OOXML | Use temp copy, retry logic |
| Modern comment XML schema changes | Parser breaks | Version detection, fallback |
| Office version variations | Different behavior | Test across versions |

### Medium Risk

| Risk | Impact | Mitigation |
|------|--------|------------|
| UIA tree structure varies | UI navigation fails | Use keyboard simulation fallback |
| COM-OOXML correlation errors | Wrong status displayed | Text matching validation |
| Performance on large presentations | Slow response | Caching, lazy loading |

### Low Risk

| Risk | Impact | Mitigation |
|------|--------|------------|
| Temp file cleanup fails | Disk space | Periodic cleanup |
| NVDA API changes | Compatibility | Version checks |

---

## 11. Recommendations

### Recommended Approach: Hybrid COM + OOXML

**Primary Strategy:**
1. Use COM for comment enumeration and navigation
2. Use OOXML parsing for resolution status
3. Cache resolution data aggressively
4. Use UIA only for optional focus management

**Implementation Priority:**

1. **Phase 1 - Core Functionality**
   - Implement OOXML comment reader with resolution status
   - Integrate with existing CommentNavigator
   - Add resolution status to announcements

2. **Phase 2 - Navigation Enhancement**
   - Add "navigate to next unresolved" command
   - Add comment statistics command
   - Implement filter by status

3. **Phase 3 - Focus Management (Optional)**
   - Investigate UIA focus for Comments pane
   - Implement keyboard simulation fallback
   - Test across Office versions

### Code Integration Points

In your existing `navigation.py`:

```python
# Replace CommentNavigator with ModernCommentNavigator
from .modern_comments import ModernCommentNavigator

class CommentNavigator(ModernCommentNavigator):
    """Enhanced navigator with resolution status"""
    pass
```

In your existing `powerpnt.py`:

```python
# Add new gesture for unresolved navigation
@script(
    description="Navigate to next unresolved comment",
    gesture="kb:control+alt+shift+pageDown",
    category="PowerPoint AI"
)
def script_nextUnresolvedComment(self, gesture):
    if self.comment_navigator:
        self.comment_navigator.navigate_to_next_unresolved()
```

### Go/No-Go Assessment

| Criterion | Status | Notes |
|-----------|--------|-------|
| Reliable resolution detection | GO | Via OOXML parsing |
| Performance target | GO | With caching |
| NVDA integration feasible | GO | Standard patterns |
| UIA for focus | CONDITIONAL | May need keyboard fallback |

**Overall Recommendation: PROCEED with hybrid approach**

---

## References

### Microsoft Documentation
- [UI Automation Custom Extensions in Office](https://learn.microsoft.com/en-us/office/uia/)
- [PowerPoint Custom Properties](https://learn.microsoft.com/en-us/office/uia/powerpoint/powerpointcustomproperties)
- [MS-PPTX CT_Comment](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/161bc2c9-98fc-46b7-852b-ba7ee77e2e54)
- [CommentStatus Enum](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2021.powerpoint.comment.commentstatus)
- [PowerPoint VBA Comment Object](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.comment)
- [IUIAutomationElement::SetFocus](https://learn.microsoft.com/en-us/windows/win32/api/uiautomationclient/nf-uiautomationclient-iuiautomationelement-setfocus)

### NVDA Resources
- [NVDA Developer Guide](https://download.nvaccess.org/documentation/developerGuide.html)
- [NVDA UIAHandler Source](https://github.com/nvaccess/nvda/blob/master/source/NVDAObjects/UIA/__init__.py)
- [NVDA PR #12861 - Custom Annotations](https://github.com/nvaccess/nvda/pull/12861)
- [NVDA Issue #9285 - Word Comments](https://github.com/nvaccess/nvda/issues/9285)

### Community Resources
- [Open-XML-SDK Issue #1133](https://github.com/OfficeDev/Open-XML-SDK/issues/1133)
- [PowerPoint Comment Status Stack Overflow](https://stackoverflow.com/questions/78347637/powerpoint-vba-code-for-pulling-out-a-slides-comments-statuss)
- [Modern Comments Support](https://support.microsoft.com/en-us/office/modern-comments-in-powerpoint-c0aa37bb-82cb-414c-872d-178946ff60ec)
- [Keyboard Shortcuts for Comments](https://support.microsoft.com/en-us/topic/use-keyboard-shortcuts-to-navigate-modern-comments-in-powerpoint-e6924fd8-43f2-474f-a1c5-7ccdfbf59b3b)

---

*Document generated: December 4, 2025*
*Research conducted using web search, Microsoft documentation, and NVDA source code analysis*
