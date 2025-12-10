# Windows Accessibility (UIA) - Expert Knowledge

## Overview

This document contains distilled knowledge for Windows UI Automation, specifically for focusing PowerPoint's Comments pane.

## Target Environment

| Component | Version |
|-----------|---------|
| Windows | Windows 11 25H2 (10.0.26200+) |
| UIA Interface | IUIAutomation6 |
| NVDA Minimum | 2025.1 |
| NVDA Tested | 2025.3.2 |
| Python | 3.11.9 (32-bit, within NVDA) |

**Note:** User's system runs Windows 11 25H2 (10.0.26200.7309), IUIAutomation6.

## Reference Files

| Topic | File | What It Contains |
|-------|------|------------------|
| Decisions | `decisions.md` | Rationale for UIA vs COM, window class targeting |
| Current Implementation | `MVP_IMPLEMENTATION_PLAN.md` | Phase-by-phase code for comment navigation |
| PowerPoint Data Access | `powerpoint-automation` agent | COM automation for reading comments, slides |
| NVDA Integration | `nvda-plugins` agent | How to integrate UIA focus in NVDA app module |
| Research | `research/` folder | Deep UIA analysis documents |

## Cross-References

**Related Agents:**
- **powerpoint-automation**: Use COM to read comment data, UIA to focus UI elements
- **nvda-plugins**: Integrate UIA focus operations within NVDA's event system

**Key Relationship:** This agent provides the UIA focus mechanism that works alongside COM data access (from powerpoint-automation) within an NVDA app module (from nvda-plugins).

## UI Automation Basics

### What is UIA?

UI Automation is Microsoft's accessibility framework that:
- Exposes UI elements to assistive technologies
- Provides programmatic access to controls
- Enables focus management and property queries

### When to Use UIA vs COM

| Use Case | Approach |
|----------|----------|
| Read PowerPoint data (slides, comments) | COM |
| Focus UI elements (panes, buttons) | UIA |
| Navigate content | COM |
| Interact with task panes | UIA |

**Our pattern:** COM for data, UIA for focus.

## UIA with comtypes

### Initialize UIA

```python
from comtypes.client import CreateObject
import comtypes.gen.UIAutomationClient as UIA

# Create automation instance
automation = CreateObject(
    "{ff48dba4-60ef-4201-aa87-54103eef594e}",
    interface=UIA.IUIAutomation
)
```

### Get Element from Window

```python
import win32gui

# Get PowerPoint window handle
hwnd = win32gui.GetForegroundWindow()

# Get UIA element
root = automation.ElementFromHandle(hwnd)
```

### Find Elements

```python
# By Name
name_cond = automation.CreatePropertyCondition(
    UIA.UIA_NamePropertyId,
    "Comments"
)
element = root.FindFirst(UIA.TreeScope_Descendants, name_cond)

# By Automation ID
id_cond = automation.CreatePropertyCondition(
    UIA.UIA_AutomationIdPropertyId,
    "CommentsPane"
)

# By Control Type
type_cond = automation.CreatePropertyCondition(
    UIA.UIA_ControlTypePropertyId,
    UIA.UIA_ListItemControlTypeId
)

# Combined conditions
combined = automation.CreateAndCondition(name_cond, type_cond)
```

### Set Focus

```python
element = find_comments_pane(root)
if element:
    element.SetFocus()
```

### Get All Matching Elements

```python
elements = root.FindAll(UIA.TreeScope_Descendants, condition)

for i in range(elements.Length):
    element = elements.GetElement(i)
    print(element.CurrentName)
```

## PowerPoint Window Classes

### Class Hierarchy

| Class | Description | UIA Status |
|-------|-------------|------------|
| PPTFrameClass | Main window | N/A |
| mdiClass | MDI container | Disabled by NVDA |
| paneClassDC | Content pane | Disabled by NVDA |
| NetUIHWND | Ribbon/task panes | Enabled |
| screenClass | Slideshow | Disabled by NVDA |

### Why NVDA Disables UIA

From NVDA source (Issue #3578):
> Microsoft's UIA implementation for PowerPoint is incomplete and "cripples existing support/hacks by other ATs"

NVDA adds these to `badUIAWindowClasses`:
- `paneClassDC`
- `mdiClass`

**Result:** Content uses COM, task panes (like Comments) use UIA.

## Comments Pane Structure

### UIA Tree (Typical)

```
NetUIHWNDElement (Comments pane)
├── Text "Comments"
├── List
│   ├── ListItem (Comment 1)
│   │   ├── Text "Author Name"
│   │   ├── Text "Comment text..."
│   │   └── Button "Reply"
│   ├── ListItem (Comment 2)
│   └── ...
└── Button "New Comment"
```

### Finding First Comment

```python
def find_first_comment(pane):
    # Find list items in the pane
    list_item_cond = automation.CreatePropertyCondition(
        UIA.UIA_ControlTypePropertyId,
        UIA.UIA_ListItemControlTypeId
    )
    items = pane.FindAll(UIA.TreeScope_Descendants, list_item_cond)

    if items and items.Length > 0:
        return items.GetElement(0)
    return None
```

### Focusing Comment at Index

```python
def focus_comment(pane, index):
    items = find_all_comments(pane)
    if items and index < items.Length:
        items.GetElement(index).SetFocus()
        return True
    return False
```

## User Identity Detection

### Windows Display Name

```python
import ctypes

def get_windows_display_name():
    GetUserNameEx = ctypes.windll.secur32.GetUserNameExW
    NameDisplay = 3  # EXTENDED_NAME_FORMAT

    # Get required buffer size
    size = ctypes.pointer(ctypes.c_ulong(0))
    GetUserNameEx(NameDisplay, None, size)

    if size.contents.value == 0:
        return None

    # Get the name
    buffer = ctypes.create_unicode_buffer(size.contents.value)
    GetUserNameEx(NameDisplay, buffer, size)

    return buffer.value if buffer.value else None
```

### Fallback Chain

```python
import os

def get_user_identity():
    # Try Windows display name first
    name = get_windows_display_name()
    if name:
        return name

    # Fallback to username
    return os.environ.get('USERNAME', 'Unknown')
```

## Common Patterns

### Retry with Delay

```python
import time

def focus_with_retry(target_func, max_attempts=3, delay=0.2):
    for attempt in range(max_attempts):
        if target_func():
            return True
        time.sleep(delay)
    return False
```

### Check if Pane Visible

```python
def is_comments_pane_visible(root):
    pane = find_comments_pane(root)
    if pane:
        # Check if actually visible
        rect = pane.CurrentBoundingRectangle
        return rect.right > rect.left and rect.bottom > rect.top
    return False
```

## Gotchas

1. **UIA elements can become stale** - Re-fetch after UI changes
2. **SetFocus may not work** - Element must be focusable
3. **FindFirst returns None** - Always check before using
4. **Timing matters** - UI needs time to update after commands
5. **NVDA intercepts focus** - Work with NVDA's focus system, not against it

## Additional Resources

For detailed research and implementation decisions:
- See `decisions.md` for rationale behind UIA vs COM choices
- See `research/PowerPoint-UIA-Research.md` for deep UIA analysis
- See `powerpoint-automation` agent for COM-based data access patterns
- See `nvda-plugins` agent for NVDA integration strategies
