# UI Automation (UIA) - Expert Knowledge

## Overview

General knowledge for Windows UI Automation in NVDA addon development. For PowerPoint-specific UIA patterns, see `powerpoint-comments/docs/experts/expert-ppt-uia.md`.

## Target Environment

| Component | Version |
|-----------|---------|
| Windows | Windows 10/11 |
| UIA Interface | IUIAutomation6 |
| Python | 3.11.9 (32-bit, within NVDA) |

## When NVDA Uses UIA

NVDA uses a decision tree to determine whether to use UIA for an element:

1. Check `goodUIAWindowClassNames` - always use UIA
2. Check `badUIAWindowClassNames` - never use UIA
3. Check app module's `shouldUseUIAInOverlay()` - app-specific override
4. Check if window has server-side UIA provider
5. Fall back to MSAA/IAccessible

### Window Class Lists

**Good UIA Classes** (always use UIA):
- `NetUIHWND` - Office ribbon and task panes
- `_WwG` - Word document windows
- `ConsoleWindowClass` - Console windows

**Bad UIA Classes** (NVDA rejects UIA):
- `paneClassDC` - PowerPoint content area (uses COM instead)
- `mdiClass` - PowerPoint MDI container
- `screenClass` - PowerPoint slideshow

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

### Get Element from Window Handle

```python
import ctypes
from ctypes import wintypes

# Get foreground window
hwnd = ctypes.windll.user32.GetForegroundWindow()

# Get UIA element
root = automation.ElementFromHandle(hwnd)
```

### Find Elements by Property

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

### Common Control Type IDs

| Control Type | ID Constant |
|--------------|-------------|
| Button | `UIA_ButtonControlTypeId` |
| Edit | `UIA_EditControlTypeId` |
| List | `UIA_ListControlTypeId` |
| ListItem | `UIA_ListItemControlTypeId` |
| Pane | `UIA_PaneControlTypeId` |
| Text | `UIA_TextControlTypeId` |
| Window | `UIA_WindowControlTypeId` |

### Common Property IDs

| Property | ID Constant |
|----------|-------------|
| Name | `UIA_NamePropertyId` |
| AutomationId | `UIA_AutomationIdPropertyId` |
| ClassName | `UIA_ClassNamePropertyId` |
| ControlType | `UIA_ControlTypePropertyId` |
| BoundingRectangle | `UIA_BoundingRectanglePropertyId` |

### Set Focus

```python
element = find_target_element(root)
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

## Tree Scope Constants

| Scope | Description |
|-------|-------------|
| `TreeScope_Element` | Just the element itself |
| `TreeScope_Children` | Direct children only |
| `TreeScope_Descendants` | All descendants (recursive) |
| `TreeScope_Subtree` | Element + all descendants |

## NVDA's UIA Infrastructure

### UIAHandler Module

Location: `source/UIAHandler/__init__.py`

Key components:
- MTA thread for COM operations
- Tree walkers for navigation
- Event handler registration

### UIA NVDAObjects

Location: `source/NVDAObjects/UIA/__init__.py`

Base class for UIA-based accessibility objects in NVDA.

### App Module Integration

```python
class AppModule(appModuleHandler.AppModule):
    # Force UIA for specific window classes
    def _get_UIAWindowClassesToAllow(self):
        return {"NetUIHWND"}

    # Block UIA for specific classes
    def _get_UIAWindowClassesToBlock(self):
        return {"CustomControl"}
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

### Check Element Visibility

```python
def is_element_visible(element):
    if element:
        rect = element.CurrentBoundingRectangle
        return rect.right > rect.left and rect.bottom > rect.top
    return False
```

### Safe Element Access

```python
def safe_get_name(element):
    try:
        return element.CurrentName if element else None
    except Exception:
        return None
```

## Gotchas

1. **UIA elements can become stale** - Re-fetch after UI changes
2. **SetFocus may not work** - Element must be focusable and visible
3. **FindFirst returns None** - Always check before using
4. **Timing matters** - UI needs time to update after commands
5. **NVDA intercepts focus** - Work with NVDA's focus system, not against it
6. **MTA vs STA** - UIA operations should run on MTA thread in NVDA

## Exploration Tools

- **Inspect.exe** - Windows SDK UIA inspection tool
- **Accessibility Insights** - Microsoft's accessibility testing tool
- **NVDA Speech Viewer** - See what NVDA announces

## Related Documentation

- **Deep Research:** `docs/research/NVDA_UIA_Deep_Research.md`
- **PowerPoint UIA:** `powerpoint-comments/docs/experts/expert-ppt-uia.md`
