# NVDA UI Automation (UIA) Deep Research

## Executive Summary

This document provides comprehensive research on how NVDA uses UI Automation (UIA), including decision logic, APIs, best practices, and specific guidance for PowerPoint plugin development.

**Key Finding for PowerPoint**: NVDA's existing PowerPoint app module explicitly DISABLES UIA in favor of COM automation because Microsoft's PowerPoint UIA implementation is incomplete. Any new plugin should follow a hybrid approach: use COM for slide content and consider UIA only for specific UI elements where COM is insufficient.

---

## Table of Contents

1. [NVDA UIA Architecture](#1-nvda-uia-architecture)
2. [When Does NVDA Use UIA vs Other APIs](#2-when-does-nvda-use-uia-vs-other-apis)
3. [UIAHandler Module Deep Dive](#3-uiahandler-module-deep-dive)
4. [Creating UIA-Based NVDA Objects](#4-creating-uia-based-nvda-objects)
5. [UIA Events in NVDA](#5-uia-events-in-nvda)
6. [UIA in App Modules](#6-uia-in-app-modules)
7. [Office-Specific UIA Handling](#7-office-specific-uia-handling)
8. [Best Practices](#8-best-practices)
9. [Common Pitfalls](#9-common-pitfalls)
10. [Code Examples](#10-code-examples)
11. [References and Resources](#11-references-and-resources)

---

## 1. NVDA UIA Architecture

### 1.1 Architecture Diagram

```
+------------------------------------------------------------------+
|                        NVDA Main Process                          |
+------------------------------------------------------------------+
|                                                                    |
|  +--------------------+     +--------------------------------+     |
|  |   Event Handler    |<----|       UIAHandler Module        |     |
|  | (eventHandler.py)  |     |  (UIAHandler/__init__.py)      |     |
|  +--------------------+     +--------------------------------+     |
|          ^                            |                            |
|          |                            v                            |
|          |                  +-------------------+                  |
|          |                  |  MTA Thread       |                  |
|          |                  |  (COM Apartment)  |                  |
|          |                  +-------------------+                  |
|          |                            |                            |
|  +-------+--------+                   v                            |
|  |  NVDAObjects   |         +-------------------+                  |
|  |   /UIA/        |<------->| IUIAutomation     |                  |
|  | (__init__.py)  |         | Client Object     |                  |
|  +----------------+         +-------------------+                  |
|          |                            |                            |
|          v                            v                            |
|  +----------------+         +-------------------+                  |
|  | App Modules    |         | Tree Walkers:     |                  |
|  | (appModules/)  |         | - baseTreeWalker  |                  |
|  +----------------+         | - windowTreeWalker|                  |
|                             +-------------------+                  |
+------------------------------------------------------------------+
                                        |
                                        v
+------------------------------------------------------------------+
|                    Windows UIA Core (UIAutomationCore.dll)        |
+------------------------------------------------------------------+
                                        |
                                        v
+------------------------------------------------------------------+
|                    Target Application (e.g., PowerPoint)          |
+------------------------------------------------------------------+
```

### 1.2 Key Components

| Component | File Location | Purpose |
|-----------|---------------|---------|
| UIAHandler | `source/UIAHandler/__init__.py` | Main UIA handler initialization, event registration |
| UIA Utils | `source/UIAHandler/utils.py` | Utility functions for element retrieval, caching |
| UIA NVDAObjects | `source/NVDAObjects/UIA/__init__.py` | Base class for UIA-based accessibility objects |
| Event Handler | `source/eventHandler.py` | Central event queue and dispatch system |
| App Modules | `source/appModules/*.py` | Application-specific handlers |

### 1.3 Initialization Flow

1. **MTA Thread Creation**: UIAHandler creates a dedicated Multi-Threaded Apartment thread to prevent UI freezing
2. **COM Object Creation**: Creates `CUIAutomation8` COM object as the UIA client
3. **Interface Version Query**: Queries for highest supported interface (IUIAutomation through IUIAutomation6)
4. **Tree Walker Setup**: Creates `windowTreeWalker` (for native windows) and `baseTreeWalker` (raw view)
5. **Event Handler Registration**: Registers global event handlers for focus, property changes

---

## 2. When Does NVDA Use UIA vs Other APIs

### 2.1 API Decision Tree

```
                        +-------------------+
                        | New Focus/Element |
                        +-------------------+
                                  |
                                  v
                    +---------------------------+
                    | Is this NVDA's own        |
                    | process?                  |
                    +---------------------------+
                           |           |
                          YES         NO
                           |           |
                           v           v
                    +--------+   +---------------------------+
                    | REJECT |   | Is window in              |
                    | UIA    |   | goodUIAWindowClassNames?  |
                    +--------+   +---------------------------+
                                       |           |
                                      YES         NO
                                       |           |
                                       v           v
                                +----------+  +---------------------------+
                                | USE UIA  |  | Does app module allow     |
                                +----------+  | this questionable window? |
                                              +---------------------------+
                                                     |           |
                                                    YES         NO
                                                     |           |
                                                     v           v
                                              +----------+  +---------------------------+
                                              | USE UIA  |  | Is window in              |
                                              +----------+  | badUIAWindowClassNames?   |
                                                            +---------------------------+
                                                                   |           |
                                                                  YES         NO
                                                                   |           |
                                                                   v           v
                                                            +----------+  +---------------------------+
                                                            | REJECT   |  | Does app module forbid    |
                                                            | UIA      |  | UIA for this window?      |
                                                            +----------+  +---------------------------+
                                                                                 |           |
                                                                                YES         NO
                                                                                 |           |
                                                                                 v           v
                                                                          +----------+  +---------------------------+
                                                                          | REJECT   |  | UiaHasServerSideProvider? |
                                                                          | UIA      |  +---------------------------+
                                                                          +----------+         |           |
                                                                                              YES         NO
                                                                                               |           |
                                                                                               v           v
                                                                                        +----------+  +-------------+
                                                                                        | USE UIA  |  | Use MSAA/   |
                                                                                        +----------+  | IAccessible |
                                                                                                      +-------------+
```

### 2.2 goodUIAWindowClassNames (Allowlist)

These window classes ALWAYS use UIA:

```python
goodUIAWindowClassNames = [
    'RAIL_WINDOW',  # Windows Defender Application Guard - always native UIA
]
```

### 2.3 badUIAWindowClassNames (Blocklist)

These window classes NEVER use UIA (fall back to MSAA):

```python
badUIAWindowClassNames = [
    "SysTreeView32",
    "WuDuiListView",
    "ComboBox",
    "msctls_progress32",
    "Edit",
    "CommonPlacesWrapperWndClass",
    "SysMonthCal32",
    "SUPERGRID",           # Outlook 2010 message list
    "RichEdit",
    "RichEdit20",
    "RICHEDIT50W",
    "SysListView32",
    "EXCEL7",
    "Button",
    "ConsoleWindowClass",  # Windows 10 has incomplete UIA for console
]
```

### 2.4 API Selection by Application Type

| Application Type | Primary API | Reason |
|------------------|-------------|--------|
| **Browsers** (Chrome, Firefox, Edge) | IAccessible2 | More battle-tested, better virtual buffer support |
| **Legacy Windows** (Desktop, Taskbar) | MSAA | Older components, mature support |
| **Modern Windows** (UWP, Windows 8+) | UIA | Native UIA support |
| **Microsoft Word** | Optional UIA | User configurable in Advanced Settings |
| **Microsoft Excel** | UIA | Rich UIA custom properties |
| **Microsoft PowerPoint** | COM | UIA explicitly disabled - incomplete implementation |
| **Java Applications** | Java Access Bridge | JAB provides better support |

### 2.5 Configuration Options

Users can configure UIA behavior in NVDA Advanced Settings:

- **Use UI Automation to access Microsoft Word document controls**: Options include "Only when necessary", "Where available", "Always"
- **Selective event registration** (Windows 11 default): Reduces performance impact
- **Browser UIA forcing**: Can be enabled via NVDA's advanced settings menu

---

## 3. UIAHandler Module Deep Dive

### 3.1 handler.clientObject - The IUIAutomation Interface

```python
# Located in UIAHandler/__init__.py
# The main COM interface for UIA operations

import UIAHandler

# Access the IUIAutomation client object
uia_client = UIAHandler.handler.clientObject

# Common operations:
# Get root element (desktop)
root = uia_client.GetRootElement()

# Get element from screen coordinates
element = uia_client.ElementFromPoint(tagPOINT(x, y))

# Get element from window handle
element = uia_client.ElementFromHandle(hwnd)

# Create tree walker with condition
condition = uia_client.CreateTrueCondition()
walker = uia_client.CreateTreeWalker(condition)
```

### 3.2 Tree Walkers

NVDA provides two pre-configured tree walkers:

```python
import UIAHandler

# Window Tree Walker - navigates to elements with native window handles
window_walker = UIAHandler.handler.windowTreeWalker

# Base Tree Walker - raw view access using RawViewWalker
base_walker = UIAHandler.handler.baseTreeWalker

# Tree walker methods:
parent = base_walker.GetParentElement(element)
first_child = base_walker.GetFirstChildElement(element)
last_child = base_walker.GetLastChildElement(element)
next_sibling = base_walker.GetNextSiblingElement(element)
prev_sibling = base_walker.GetPreviousSiblingElement(element)

# With caching (preferred for performance):
parent = base_walker.GetParentElementBuildCache(element, cache_request)
```

### 3.3 Element Caching Strategies

```python
import UIAHandler

# Create a cache request
cache_request = UIAHandler.handler.clientObject.CreateCacheRequest()

# Add properties to cache
cache_request.AddProperty(UIAHandler.UIA_NamePropertyId)
cache_request.AddProperty(UIAHandler.UIA_ControlTypePropertyId)
cache_request.AddProperty(UIAHandler.UIA_AutomationIdPropertyId)

# Set caching scope
cache_request.TreeScope = UIAHandler.TreeScope_Element | UIAHandler.TreeScope_Children

# Fetch element with cache
element = UIAHandler.handler.clientObject.ElementFromHandleBuildCache(hwnd, cache_request)

# Access cached properties (faster than GetCurrentPropertyValue)
name = element.GetCachedPropertyValue(UIAHandler.UIA_NamePropertyId)
```

### 3.4 Key UIAHandler Methods

| Method | Purpose |
|--------|---------|
| `isUIAWindow(hwnd)` | Determines if a window natively supports UIA (cached 0.5s) |
| `isNativeUIAElement(element)` | Validates element is genuinely UIA-native vs proxied from MSAA |
| `getNearestWindowHandle(element)` | Walks UIA tree upward to find valid window handle |
| `IUIAFocusChangedEventHandler_HandleFocusChangedEvent()` | Handles focus change events |
| `IUIAPropertyChangedEventHandler_HandlePropertyChangedEvent()` | Handles property change events |

---

## 4. Creating UIA-Based NVDA Objects

### 4.1 UIA NVDAObject Base Class

```python
# Located in NVDAObjects/UIA/__init__.py

from NVDAObjects.UIA import UIA

class MyCustomUIAObject(UIA):
    """Custom UIA object for specific control."""

    def _get_name(self):
        """Override name retrieval."""
        # Use cached property if available
        name = self._getUIACacheablePropertyValue(UIAHandler.UIA_NamePropertyId)
        if not name:
            # Fallback logic
            name = "Unknown"
        return name

    def _get_role(self):
        """Override role detection."""
        # Map UIA control type to NVDA role
        control_type = self._getUIACacheablePropertyValue(UIAHandler.UIA_ControlTypePropertyId)
        return UIAHandler.UIAControlTypesToNVDARoles.get(control_type, controlTypes.Role.UNKNOWN)

    def _get_states(self):
        """Override state detection."""
        states = super()._get_states()
        # Add custom state logic
        if self._getUIACacheablePropertyValue(UIAHandler.UIA_IsEnabledPropertyId):
            states.discard(controlTypes.State.UNAVAILABLE)
        return states
```

### 4.2 Role Mapping (UIA ControlType to NVDA Role)

NVDA maps UIA control types to NVDA roles in `UIAHandler.UIAControlTypesToNVDARoles`:

```python
# Key mappings (partial list):
UIAControlTypesToNVDARoles = {
    UIAHandler.UIA_ButtonControlTypeId: controlTypes.Role.BUTTON,
    UIAHandler.UIA_EditControlTypeId: controlTypes.Role.EDITABLETEXT,
    UIAHandler.UIA_TextControlTypeId: controlTypes.Role.STATICTEXT,
    UIAHandler.UIA_ListControlTypeId: controlTypes.Role.LIST,
    UIAHandler.UIA_ListItemControlTypeId: controlTypes.Role.LISTITEM,
    UIAHandler.UIA_TreeControlTypeId: controlTypes.Role.TREEVIEW,
    UIAHandler.UIA_TreeItemControlTypeId: controlTypes.Role.TREEVIEWITEM,
    UIAHandler.UIA_DocumentControlTypeId: controlTypes.Role.DOCUMENT,
    UIAHandler.UIA_PaneControlTypeId: controlTypes.Role.PANE,
    UIAHandler.UIA_WindowControlTypeId: controlTypes.Role.WINDOW,
    # ... many more
}
```

### 4.3 Pattern Retrieval

```python
class MyUIAObject(UIA):
    """Example showing pattern retrieval."""

    def _get_UIAInvokePattern(self):
        """Lazy-load Invoke pattern."""
        return self._getUIAPattern(UIAHandler.UIA_InvokePatternId,
                                   UIAHandler.IUIAutomationInvokePattern)

    def _get_UIAValuePattern(self):
        """Lazy-load Value pattern."""
        return self._getUIAPattern(UIAHandler.UIA_ValuePatternId,
                                   UIAHandler.IUIAutomationValuePattern)

    def _get_UIATogglePattern(self):
        """Lazy-load Toggle pattern."""
        return self._getUIAPattern(UIAHandler.UIA_TogglePatternId,
                                   UIAHandler.IUIAutomationTogglePattern)

    def doAction(self):
        """Invoke the control."""
        invoke_pattern = self.UIAInvokePattern
        if invoke_pattern:
            invoke_pattern.Invoke()

    def _get_value(self):
        """Get control value."""
        value_pattern = self.UIAValuePattern
        if value_pattern:
            return value_pattern.CurrentValue
        return ""
```

### 4.4 UIA Overlay Classes

Used to customize behavior for specific controls without full class replacement:

**NOTE:** This example shows a NEW app module (no built-in support to extend).
For apps WITH built-in support (like PowerPoint), inherit from the built-in
AppModule class instead. See `decisions.md` Decision 6.

```python
# In appModules/myapp.py (for app WITHOUT built-in support)

class AppModule(appModuleHandler.AppModule):

    def chooseNVDAObjectOverlayClasses(self, obj, clsList):
        """Add overlay classes based on element properties."""
        if isinstance(obj, UIA):
            automation_id = obj.UIAAutomationId
            if automation_id == "MySpecialControl":
                clsList.insert(0, MySpecialControlOverlay)

class MySpecialControlOverlay(UIA):
    """Overlay that adds special handling."""

    def _get_name(self):
        # Custom name logic
        return "Special: " + super()._get_name()
```

---

## 5. UIA Events in NVDA

### 5.1 Event Types Supported

| Event Type | UIA Event ID | NVDA Event Name | Description |
|------------|--------------|-----------------|-------------|
| Focus Changed | N/A (special) | `gainFocus` | Element received focus |
| Name Changed | UIA_NamePropertyId | `nameChange` | Name property changed |
| Value Changed | UIA_ValueValuePropertyId | `valueChange` | Value changed |
| Selection Changed | UIA_SelectionItem_* | `stateChange` | Selection state changed |
| Notification | UIA_NotificationEventId | `UIA_notification` | App notification |
| System Alert | UIA_SystemAlertEventId | `UIA_systemAlert` | System alert |

### 5.2 Event Subscription Methods

```python
import UIAHandler

# Global event subscription (listens everywhere)
UIAHandler.handler.clientObject.AddFocusChangedEventHandler(
    cache_request,
    event_handler
)

# Scoped event subscription (specific subtree)
UIAHandler.handler.clientObject.AddAutomationEventHandler(
    UIAHandler.UIA_Invoke_InvokedEventId,
    root_element,
    UIAHandler.TreeScope_Subtree,
    cache_request,
    event_handler
)

# Property changed subscription
UIAHandler.handler.clientObject.AddPropertyChangedEventHandler(
    root_element,
    UIAHandler.TreeScope_Subtree,
    cache_request,
    event_handler,
    (UIAHandler.UIA_NamePropertyId, UIAHandler.UIA_ValueValuePropertyId)
)
```

### 5.3 Event Coalescing and Performance

NVDA implements sophisticated event coalescing to avoid flooding:

- **PropertyChange events**: Coalesced if runtimeIDs match and same property
- **Focus events**: Do NOT support coalescing (processed immediately)
- **Coalescing delay**: ~30ms window for collecting duplicate events
- **Selective registration** (Windows 11): Only registers events for focused object and ancestors

### 5.4 Event Handler Best Practices

```python
def handleEvent(self, eventID, sender, args):
    """Event handler pattern."""
    try:
        # Validate the element still exists
        if not sender:
            return

        # Quick checks before expensive operations
        try:
            runtime_id = sender.GetRuntimeId()
        except COMError:
            return  # Element no longer valid

        # Queue event for NVDA's main thread
        eventHandler.queueEvent("myEvent", obj)

    except Exception:
        log.exception("Error handling event")
```

---

## 6. UIA in App Modules

### 6.1 Accessing UIAHandler from App Module

**NOTE:** Examples in this section show NEW app modules (no built-in support).
For apps WITH built-in support (like PowerPoint), see `decisions.md` Decision 6.

```python
# appModules/myapp.py (for app WITHOUT built-in support)

import appModuleHandler
import UIAHandler
from NVDAObjects.UIA import UIA
import api

class AppModule(appModuleHandler.AppModule):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Store reference to UIA handler
        self._uiaHandler = UIAHandler.handler

    def event_gainFocus(self, obj, nextHandler):
        """Handle focus events with UIA checks."""
        if isinstance(obj, UIA):
            # Access UIA element directly
            uia_element = obj.UIAElement
            automation_id = uia_element.GetCurrentPropertyValue(
                UIAHandler.UIA_AutomationIdPropertyId
            )
            # Custom logic based on automation ID
        nextHandler()
```

### 6.2 Finding Elements Programmatically

```python
import UIAHandler
from ctypes import POINTER
from comtypes import COMError

def findElementByAutomationId(root_element, automation_id):
    """Find element by automation ID."""
    condition = UIAHandler.handler.clientObject.CreatePropertyCondition(
        UIAHandler.UIA_AutomationIdPropertyId,
        automation_id
    )
    try:
        element = root_element.FindFirst(
            UIAHandler.TreeScope_Descendants,
            condition
        )
        return element
    except COMError:
        return None

def findAllButtons(root_element):
    """Find all button elements."""
    condition = UIAHandler.handler.clientObject.CreatePropertyCondition(
        UIAHandler.UIA_ControlTypePropertyId,
        UIAHandler.UIA_ButtonControlTypeId
    )
    try:
        elements = root_element.FindAll(
            UIAHandler.TreeScope_Descendants,
            condition
        )
        return elements
    except COMError:
        return None

def findElementByMultipleConditions(root_element, name, control_type):
    """Find element matching multiple conditions."""
    name_condition = UIAHandler.handler.clientObject.CreatePropertyCondition(
        UIAHandler.UIA_NamePropertyId,
        name
    )
    type_condition = UIAHandler.handler.clientObject.CreatePropertyCondition(
        UIAHandler.UIA_ControlTypePropertyId,
        control_type
    )
    combined = UIAHandler.handler.clientObject.CreateAndCondition(
        name_condition,
        type_condition
    )
    try:
        return root_element.FindFirst(
            UIAHandler.TreeScope_Descendants,
            combined
        )
    except COMError:
        return None
```

### 6.3 Setting Focus via UIA

```python
def setFocusToElement(element):
    """Set focus to a UIA element."""
    try:
        element.SetFocus()
        return True
    except COMError:
        return False

def focusElementByAutomationId(root_hwnd, automation_id):
    """Find and focus an element by automation ID."""
    root = UIAHandler.handler.clientObject.ElementFromHandle(root_hwnd)
    element = findElementByAutomationId(root, automation_id)
    if element:
        return setFocusToElement(element)
    return False
```

### 6.4 Invoking Patterns

```python
def invokeButton(element):
    """Invoke a button element."""
    try:
        pattern = element.GetCurrentPattern(UIAHandler.UIA_InvokePatternId)
        if pattern:
            invoke = pattern.QueryInterface(UIAHandler.IUIAutomationInvokePattern)
            invoke.Invoke()
            return True
    except COMError:
        pass
    return False

def toggleCheckbox(element):
    """Toggle a checkbox element."""
    try:
        pattern = element.GetCurrentPattern(UIAHandler.UIA_TogglePatternId)
        if pattern:
            toggle = pattern.QueryInterface(UIAHandler.IUIAutomationTogglePattern)
            toggle.Toggle()
            return True
    except COMError:
        pass
    return False

def setValue(element, value):
    """Set value on an element with ValuePattern."""
    try:
        pattern = element.GetCurrentPattern(UIAHandler.UIA_ValuePatternId)
        if pattern:
            value_pattern = pattern.QueryInterface(UIAHandler.IUIAutomationValuePattern)
            value_pattern.SetValue(value)
            return True
    except COMError:
        pass
    return False
```

### 6.5 Reading Properties

```python
def getElementInfo(element):
    """Get comprehensive element information."""
    info = {}
    try:
        info['name'] = element.GetCurrentPropertyValue(UIAHandler.UIA_NamePropertyId)
        info['automation_id'] = element.GetCurrentPropertyValue(UIAHandler.UIA_AutomationIdPropertyId)
        info['control_type'] = element.GetCurrentPropertyValue(UIAHandler.UIA_ControlTypePropertyId)
        info['class_name'] = element.GetCurrentPropertyValue(UIAHandler.UIA_ClassNamePropertyId)
        info['is_enabled'] = element.GetCurrentPropertyValue(UIAHandler.UIA_IsEnabledPropertyId)
        info['has_keyboard_focus'] = element.GetCurrentPropertyValue(UIAHandler.UIA_HasKeyboardFocusPropertyId)
        info['bounding_rect'] = element.GetCurrentPropertyValue(UIAHandler.UIA_BoundingRectanglePropertyId)
    except COMError as e:
        info['error'] = str(e)
    return info
```

---

## 7. Office-Specific UIA Handling

### 7.1 PowerPoint: UIA is DISABLED

**Critical**: NVDA's PowerPoint app module explicitly disables UIA:

```python
# From appModules/powerpnt.py
# "PowerPoint 2013 implements UIA support for its slides etc on an mdiClass window.
# However its far from complete. We must disable it in order to fall back to our own code."

badUIAWindowClasses = [
    'paneClassDC',
    'mdiClass',
    'screenClass',
]
```

**Reason**: PowerPoint's UIA implementation is incomplete. NVDA uses COM automation instead.

### 7.2 Excel: UIA with Custom Properties

Excel uses UIA extensively with custom properties:

```python
# Custom property GUIDs registered by Excel
class ExcelCustomProperties:
    CELL_FORMULA = GUID('{...}')
    NUMBER_FORMAT = GUID('{...}')
    DATA_VALIDATION = GUID('{...}')
    HAS_CONDITIONAL_FORMATTING = GUID('{...}')
    COMMENT_THREAD = GUID('{...}')

# Accessing custom properties
custom_annotation_types = element.GetCurrentPropertyValue(
    UIAHandler.UIA_AnnotationTypesPropertyId
)
```

### 7.3 Word: Optional UIA

Word's UIA support is configurable:

```python
# User can select in Advanced Settings:
# - "Only when necessary" (default)
# - "Where available"
# - "Always"

# Word-specific text handling
class WordDocumentTextInfo(UIATextInfo):
    """Special handling for Word documents."""

    def _get_text(self):
        text = super()._get_text()
        # Strip end-of-row markers
        text = text.replace('\x07', '')
        # Convert vertical tabs
        text = text.replace('\x0b', '\r')
        return text
```

### 7.4 NetUIHWND Handling (Office Ribbon)

The Office ribbon uses NetUIHWND windows with known issues:

```python
# From NVDA issues:
# "UIA implementation for MS Office ribbons does not fire a focus event
# the first time the ribbon is activated after starting NVDA"

# Workaround: Ignore UIA for NetUIHWND that is descendant of MsoCommandBar
def isGoodUIAWindow(hwnd):
    """Check if this NetUIHWND should use UIA."""
    class_name = winUser.getClassName(hwnd)
    if class_name == "NetUIHWND":
        parent = winUser.getAncestor(hwnd, winUser.GA_PARENT)
        parent_class = winUser.getClassName(parent)
        if parent_class == "MsoCommandBar":
            return False
    return True
```

### 7.5 Task Panes

```python
# Task panes may use UIA but require special handling
# Access via focus events and object navigation rather than direct UIA queries
```

---

## 8. Best Practices

### 8.1 When to Use UIA vs COM in Office Apps

| Scenario | Recommended API | Reason |
|----------|-----------------|--------|
| PowerPoint slide content | COM | UIA incomplete |
| PowerPoint ribbon | UIA (careful) | Focus issues |
| Excel cells | UIA | Rich custom properties |
| Word documents | UIA (optional) | User configurable |
| Office task panes | UIA | Modern UI |
| Office dialogs | Mixed | Depends on control type |

### 8.2 Performance Optimization

1. **Use Caching**:
```python
# Create cache request with needed properties
cache_request = UIAHandler.handler.clientObject.CreateCacheRequest()
cache_request.AddProperty(UIAHandler.UIA_NamePropertyId)
cache_request.AddProperty(UIAHandler.UIA_ControlTypePropertyId)

# Fetch with cache
element = root.FindFirstBuildCache(scope, condition, cache_request)
```

2. **Limit Tree Scope**:
```python
# Bad: Search entire tree
element = root.FindFirst(TreeScope_Descendants, condition)

# Good: Limit scope
element = root.FindFirst(TreeScope_Children, condition)
```

3. **Cache RuntimeIDs for Comparison**:
```python
# Cache runtime IDs instead of comparing element objects
cached_runtime_id = element.GetRuntimeId()
# Later...
if element.GetRuntimeId() == cached_runtime_id:
    # Same element
```

### 8.3 Error Handling

```python
from comtypes import COMError

def safeGetProperty(element, property_id):
    """Safely get a UIA property."""
    try:
        return element.GetCurrentPropertyValue(property_id)
    except COMError:
        return None
    except AttributeError:
        return None

def safeGetPattern(element, pattern_id, interface):
    """Safely get a UIA pattern."""
    try:
        pattern = element.GetCurrentPattern(pattern_id)
        if pattern:
            return pattern.QueryInterface(interface)
    except COMError:
        pass
    except AttributeError:
        pass
    return None
```

### 8.4 Threading Considerations

```python
# UIA operations should be performed from the correct apartment
# NVDA uses an MTA thread for UIA handler

import threading
from queue import Queue

class UIAWorker:
    def __init__(self):
        self._queue = Queue()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def _run(self):
        import comtypes
        comtypes.CoInitializeEx(comtypes.COINIT_MULTITHREADED)
        while True:
            func, args, result_queue = self._queue.get()
            try:
                result = func(*args)
                result_queue.put(('success', result))
            except Exception as e:
                result_queue.put(('error', e))

    def execute(self, func, *args):
        result_queue = Queue()
        self._queue.put((func, args, result_queue))
        status, result = result_queue.get()
        if status == 'error':
            raise result
        return result
```

### 8.5 Best Practices Checklist

- [ ] Check if UIA is appropriate for the target control/window
- [ ] Use caching for frequently accessed properties
- [ ] Handle COMError exceptions gracefully
- [ ] Limit tree traversal scope
- [ ] Validate elements before use (may become invalid)
- [ ] Consider threading/apartment requirements
- [ ] Test with actual screen reader workflows
- [ ] Document any UIA limitations found

---

## 9. Common Pitfalls

### 9.1 Known NVDA UIA Issues

| Issue | Description | Workaround |
|-------|-------------|------------|
| PowerPoint UIA incomplete | Slides not accessible via UIA | Use COM automation |
| NetUIHWND focus issues | Ribbon doesn't fire focus first time | Disable UIA for MsoCommandBar descendants |
| Chrome performance | UIA events cause freezes | Use IAccessible2 instead |
| Element lifecycle | Elements become invalid after changes | Re-query elements, handle COMError |
| Console UIA incomplete | Windows console has limited UIA | Listed in badUIAWindowClassNames |

### 9.2 Office UIA Quirks

```python
# Issue: Office UIA doesn't always report unavailable items correctly
# Workaround: Check IsEnabled property explicitly

# Issue: Excel merged cells return as single element
# Workaround: Calculate ranges manually

# Issue: Word selection requires visible range
# Workaround: Scroll range into view before setting selection

def ensureRangeVisible(text_range):
    """Ensure text range is visible before selection."""
    try:
        text_range.ScrollIntoView(True)
    except COMError:
        pass
```

### 9.3 Race Conditions

```python
# Problem: Element might change between query and use
# Solution: Always handle COMError, re-query if needed

def robustElementAccess(root, automation_id, max_retries=3):
    """Access element with retry logic."""
    for attempt in range(max_retries):
        try:
            element = findElementByAutomationId(root, automation_id)
            if element:
                # Verify element is still valid
                _ = element.GetCurrentPropertyValue(UIAHandler.UIA_NamePropertyId)
                return element
        except COMError:
            if attempt < max_retries - 1:
                import time
                time.sleep(0.1)
    return None
```

### 9.4 Memory Leaks

```python
# Problem: COM objects not properly released
# Solution: Use context managers or explicit release

class UIAElementContext:
    """Context manager for UIA elements."""

    def __init__(self, element):
        self.element = element

    def __enter__(self):
        return self.element

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.element:
            try:
                import comtypes
                comtypes.Release(self.element)
            except:
                pass
        return False
```

### 9.5 Element Lifecycle Issues

```python
# Elements become invalid after:
# - Window closes
# - Control is destroyed
# - Document changes
# - Parent element changes

def isElementValid(element):
    """Check if element is still valid."""
    try:
        # Try to access a basic property
        _ = element.GetRuntimeId()
        return True
    except (COMError, AttributeError):
        return False
```

---

## 10. Code Examples

### 10.1 Complete App Module with UIA

**NOTE:** This example shows a NEW app module (no built-in support to extend).
For apps WITH built-in support (like PowerPoint), inherit from the built-in
AppModule class instead. See `decisions.md` Decision 6.

```python
# appModules/myapp.py (for app WITHOUT built-in support)

import appModuleHandler
import UIAHandler
from NVDAObjects.UIA import UIA
import controlTypes
import api
import ui
from logHandler import log
from comtypes import COMError

class AppModule(appModuleHandler.AppModule):
    """App module demonstrating UIA usage (no built-in to extend)."""

    # Window classes where UIA should be disabled
    badUIAWindowClasses = ['LegacyControl']

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._cachedElements = {}

    def isBadUIAWindow(self, hwnd):
        """Override to control UIA usage per window."""
        import winUser
        class_name = winUser.getClassName(hwnd)
        if class_name in self.badUIAWindowClasses:
            return True
        return False

    def chooseNVDAObjectOverlayClasses(self, obj, clsList):
        """Add custom overlay classes for specific controls."""
        if isinstance(obj, UIA):
            try:
                automation_id = obj.UIAElement.GetCurrentPropertyValue(
                    UIAHandler.UIA_AutomationIdPropertyId
                )
                if automation_id == "SpecialButton":
                    clsList.insert(0, SpecialButtonOverlay)
                elif automation_id == "CustomList":
                    clsList.insert(0, CustomListOverlay)
            except COMError:
                pass

    def event_gainFocus(self, obj, nextHandler):
        """Custom focus handling."""
        if isinstance(obj, UIA):
            try:
                control_type = obj.UIAElement.GetCurrentPropertyValue(
                    UIAHandler.UIA_ControlTypePropertyId
                )
                # Log focus for debugging
                log.debug(f"Focus on control type: {control_type}")
            except COMError:
                pass
        nextHandler()

    def findUIAElement(self, automation_id):
        """Find a UIA element by automation ID."""
        try:
            # Get foreground window
            fg = api.getForegroundObject()
            if not hasattr(fg, 'UIAElement'):
                return None

            root = fg.UIAElement
            condition = UIAHandler.handler.clientObject.CreatePropertyCondition(
                UIAHandler.UIA_AutomationIdPropertyId,
                automation_id
            )
            return root.FindFirst(
                UIAHandler.TreeScope_Descendants,
                condition
            )
        except COMError:
            return None

    def script_activateSpecialButton(self, gesture):
        """Script to activate a specific button."""
        element = self.findUIAElement("SpecialButton")
        if element:
            try:
                pattern = element.GetCurrentPattern(UIAHandler.UIA_InvokePatternId)
                if pattern:
                    invoke = pattern.QueryInterface(UIAHandler.IUIAutomationInvokePattern)
                    invoke.Invoke()
                    ui.message("Activated special button")
                    return
            except COMError:
                pass
        ui.message("Special button not found")

    __gestures = {
        "kb:NVDA+shift+s": "activateSpecialButton",
    }


class SpecialButtonOverlay(UIA):
    """Overlay for special button control."""

    def _get_name(self):
        """Custom name with prefix."""
        original = super()._get_name()
        return f"Special: {original}"

    def _get_description(self):
        """Add helpful description."""
        return "Press Enter to activate this special function"


class CustomListOverlay(UIA):
    """Overlay for custom list control."""

    def _get_role(self):
        """Force list role."""
        return controlTypes.Role.LIST

    def _get_positionInfo(self):
        """Provide position information."""
        try:
            # Try to get item index from UIA
            element = self.UIAElement
            selection_pattern = element.GetCurrentPattern(
                UIAHandler.UIA_SelectionItemPatternId
            )
            if selection_pattern:
                si = selection_pattern.QueryInterface(
                    UIAHandler.IUIAutomationSelectionItemPattern
                )
                # Get parent for total count
                container = si.CurrentSelectionContainer
                if container:
                    # Count items
                    children = container.FindAll(
                        UIAHandler.TreeScope_Children,
                        UIAHandler.handler.clientObject.CreateTrueCondition()
                    )
                    total = children.Length
                    # Determine current index
                    # (simplified - would need actual implementation)
                    return {'indexInGroup': 1, 'similarItemsInGroup': total}
        except COMError:
            pass
        return super()._get_positionInfo()
```

### 10.2 UIA Helper Module

```python
# lib/uia_helpers.py

"""Helper functions for UIA operations in NVDA plugins."""

import UIAHandler
from comtypes import COMError
from logHandler import log

class UIAHelpers:
    """Collection of UIA utility functions."""

    @staticmethod
    def get_client():
        """Get the IUIAutomation client object."""
        return UIAHandler.handler.clientObject

    @staticmethod
    def create_property_condition(property_id, value):
        """Create a property condition."""
        return UIAHelpers.get_client().CreatePropertyCondition(property_id, value)

    @staticmethod
    def create_and_condition(*conditions):
        """Create an AND condition from multiple conditions."""
        client = UIAHelpers.get_client()
        result = conditions[0]
        for cond in conditions[1:]:
            result = client.CreateAndCondition(result, cond)
        return result

    @staticmethod
    def create_or_condition(*conditions):
        """Create an OR condition from multiple conditions."""
        client = UIAHelpers.get_client()
        result = conditions[0]
        for cond in conditions[1:]:
            result = client.CreateOrCondition(result, cond)
        return result

    @staticmethod
    def find_element(root, automation_id=None, name=None, control_type=None):
        """Find an element by various criteria."""
        conditions = []

        if automation_id:
            conditions.append(UIAHelpers.create_property_condition(
                UIAHandler.UIA_AutomationIdPropertyId, automation_id
            ))
        if name:
            conditions.append(UIAHelpers.create_property_condition(
                UIAHandler.UIA_NamePropertyId, name
            ))
        if control_type:
            conditions.append(UIAHelpers.create_property_condition(
                UIAHandler.UIA_ControlTypePropertyId, control_type
            ))

        if not conditions:
            return None

        condition = conditions[0] if len(conditions) == 1 else \
                    UIAHelpers.create_and_condition(*conditions)

        try:
            return root.FindFirst(UIAHandler.TreeScope_Descendants, condition)
        except COMError:
            return None

    @staticmethod
    def find_all_elements(root, control_type):
        """Find all elements of a given control type."""
        condition = UIAHelpers.create_property_condition(
            UIAHandler.UIA_ControlTypePropertyId, control_type
        )
        try:
            return root.FindAll(UIAHandler.TreeScope_Descendants, condition)
        except COMError:
            return None

    @staticmethod
    def get_property_safe(element, property_id, default=None):
        """Safely get a property value."""
        try:
            value = element.GetCurrentPropertyValue(property_id)
            return value if value is not None else default
        except COMError:
            return default

    @staticmethod
    def invoke_element(element):
        """Invoke an element if it supports InvokePattern."""
        try:
            pattern = element.GetCurrentPattern(UIAHandler.UIA_InvokePatternId)
            if pattern:
                invoke = pattern.QueryInterface(UIAHandler.IUIAutomationInvokePattern)
                invoke.Invoke()
                return True
        except COMError as e:
            log.debug(f"Failed to invoke element: {e}")
        return False

    @staticmethod
    def set_value(element, value):
        """Set value on an element if it supports ValuePattern."""
        try:
            pattern = element.GetCurrentPattern(UIAHandler.UIA_ValuePatternId)
            if pattern:
                value_pattern = pattern.QueryInterface(UIAHandler.IUIAutomationValuePattern)
                value_pattern.SetValue(value)
                return True
        except COMError as e:
            log.debug(f"Failed to set value: {e}")
        return False

    @staticmethod
    def get_value(element):
        """Get value from an element if it supports ValuePattern."""
        try:
            pattern = element.GetCurrentPattern(UIAHandler.UIA_ValuePatternId)
            if pattern:
                value_pattern = pattern.QueryInterface(UIAHandler.IUIAutomationValuePattern)
                return value_pattern.CurrentValue
        except COMError:
            pass
        return None

    @staticmethod
    def set_focus(element):
        """Set focus to an element."""
        try:
            element.SetFocus()
            return True
        except COMError:
            return False

    @staticmethod
    def is_element_valid(element):
        """Check if an element is still valid."""
        try:
            _ = element.GetRuntimeId()
            return True
        except (COMError, AttributeError):
            return False

    @staticmethod
    def get_element_info(element):
        """Get comprehensive element information for debugging."""
        info = {}
        properties = [
            ('name', UIAHandler.UIA_NamePropertyId),
            ('automation_id', UIAHandler.UIA_AutomationIdPropertyId),
            ('control_type', UIAHandler.UIA_ControlTypePropertyId),
            ('class_name', UIAHandler.UIA_ClassNamePropertyId),
            ('is_enabled', UIAHandler.UIA_IsEnabledPropertyId),
            ('has_keyboard_focus', UIAHandler.UIA_HasKeyboardFocusPropertyId),
        ]
        for name, prop_id in properties:
            info[name] = UIAHelpers.get_property_safe(element, prop_id)
        return info

    @staticmethod
    def walk_children(element, callback):
        """Walk all children of an element, calling callback for each."""
        walker = UIAHandler.handler.baseTreeWalker
        try:
            child = walker.GetFirstChildElement(element)
            while child:
                callback(child)
                child = walker.GetNextSiblingElement(child)
        except COMError:
            pass

    @staticmethod
    def walk_ancestors(element, callback):
        """Walk ancestors of an element, calling callback for each."""
        walker = UIAHandler.handler.baseTreeWalker
        try:
            parent = walker.GetParentElement(element)
            while parent:
                if callback(parent) is False:
                    break
                parent = walker.GetParentElement(parent)
        except COMError:
            pass
```

### 10.3 PowerPoint Plugin with Hybrid Approach

**NOTE:** This example shows the correct pattern for extending built-in PowerPoint support.
See `decisions.md` Decision 6 for full details.

```python
# appModules/powerpnt.py (extends built-in PowerPoint support)

"""
Enhanced PowerPoint plugin using hybrid COM/UIA approach.
Uses COM for slide content (where UIA is incomplete)
and UIA for modern UI elements like task panes.
"""

import UIAHandler
from NVDAObjects.UIA import UIA
from NVDAObjects.IAccessible import IAccessible
import api
import ui
import controlTypes
from logHandler import log
from comtypes import COMError
import winUser

# Import built-in PowerPoint module to extend it
# Pattern reference: decisions.md Decision 6
from nvdaBuiltin.appModules.powerpnt import AppModule as BuiltinPowerPointAppModule

class AppModule(BuiltinPowerPointAppModule):
    """Enhanced PowerPoint app module with selective UIA usage."""

    # Task pane classes that CAN use UIA
    taskPaneClasses = ['NetUIHWND']

    # Main PowerPoint classes that should NOT use UIA
    badUIAWindowClasses = BuiltinPowerPointAppModule.badUIAWindowClasses + [
        'paneClassDC',
        'mdiClass',
        'screenClass',
    ]

    def isBadUIAWindow(self, hwnd):
        """
        Selectively allow UIA for task panes.
        """
        class_name = winUser.getClassName(hwnd)

        # Allow UIA for task panes (they work better with UIA)
        if class_name in self.taskPaneClasses:
            parent_class = winUser.getClassName(
                winUser.getAncestor(hwnd, winUser.GA_PARENT)
            )
            # Only if NOT part of the ribbon
            if parent_class != "MsoCommandBar":
                return False  # Allow UIA

        # Block UIA for main PowerPoint windows
        if class_name in self.badUIAWindowClasses:
            return True

        return super().isBadUIAWindow(hwnd)

    def chooseNVDAObjectOverlayClasses(self, obj, clsList):
        """Add overlay classes for task pane elements."""
        if isinstance(obj, UIA):
            try:
                automation_id = obj.UIAElement.GetCurrentPropertyValue(
                    UIAHandler.UIA_AutomationIdPropertyId
                )
                # Handle Accessibility Checker task pane
                if "AccessibilityChecker" in str(automation_id):
                    clsList.insert(0, AccessibilityCheckerOverlay)
                # Handle Design Ideas task pane
                elif "DesignIdeas" in str(automation_id):
                    clsList.insert(0, DesignIdeasOverlay)
            except COMError:
                pass

        # Let base class add its overlays
        super().chooseNVDAObjectOverlayClasses(obj, clsList)

    def findTaskPaneElement(self, automation_id_contains):
        """Find an element in a task pane by partial automation ID."""
        try:
            fg = api.getForegroundObject()
            # Walk windows to find task pane
            for child in fg.children:
                if isinstance(child, UIA):
                    try:
                        aid = child.UIAElement.GetCurrentPropertyValue(
                            UIAHandler.UIA_AutomationIdPropertyId
                        )
                        if aid and automation_id_contains in aid:
                            return child
                    except COMError:
                        continue
        except Exception as e:
            log.debug(f"Error finding task pane element: {e}")
        return None

    def script_announceAccessibilityIssues(self, gesture):
        """Read accessibility checker results."""
        element = self.findTaskPaneElement("AccessibilityChecker")
        if element:
            try:
                # Navigate to issues list in task pane
                issues = []
                for child in element.children:
                    if child.role == controlTypes.Role.LISTITEM:
                        issues.append(child.name)

                if issues:
                    ui.message(f"Found {len(issues)} accessibility issues: " +
                              ", ".join(issues[:5]))
                else:
                    ui.message("No accessibility issues found")
                return
            except Exception as e:
                log.debug(f"Error reading accessibility issues: {e}")

        ui.message("Accessibility checker not open")

    __gestures = {
        "kb:NVDA+shift+a": "announceAccessibilityIssues",
    }


class AccessibilityCheckerOverlay(UIA):
    """Overlay for Accessibility Checker task pane elements."""

    def _get_name(self):
        """Provide descriptive names for checker elements."""
        name = super()._get_name()
        if not name:
            # Try to get from child text elements
            for child in self.children:
                if child.role == controlTypes.Role.STATICTEXT:
                    return child.name
        return name or "Accessibility item"


class DesignIdeasOverlay(UIA):
    """Overlay for Design Ideas task pane elements."""

    def _get_name(self):
        """Provide descriptive names for design ideas."""
        name = super()._get_name()
        return f"Design suggestion: {name}" if name else "Design suggestion"
```

---

## 11. References and Resources

### 11.1 NVDA Source Code References

- [UIAHandler/__init__.py](https://github.com/nvaccess/nvda/blob/master/source/UIAHandler/__init__.py) - Main UIA handler
- [UIAHandler/utils.py](https://github.com/nvaccess/nvda/blob/master/source/UIAHandler/utils.py) - UIA utilities
- [NVDAObjects/UIA/__init__.py](https://github.com/nvaccess/nvda/blob/master/source/NVDAObjects/UIA/__init__.py) - UIA NVDA objects
- [NVDAObjects/UIA/excel.py](https://github.com/nvaccess/nvda/blob/master/source/NVDAObjects/UIA/excel.py) - Excel UIA handling
- [NVDAObjects/UIA/wordDocument.py](https://github.com/nvaccess/nvda/blob/master/source/NVDAObjects/UIA/wordDocument.py) - Word UIA handling
- [appModules/powerpnt.py](https://github.com/nvaccess/nvda/blob/master/source/appModules/powerpnt.py) - PowerPoint app module

### 11.2 Relevant NVDA GitHub Issues

- [Issue #4207 - Alt key focus issues with MS Office ribbons](https://github.com/nvaccess/nvda/issues/4207)
- [Issue #6437 - UIA event tracking for plugins](https://github.com/nvaccess/nvda/issues/6437)
- [Issue #7409 - Switch to UIA for Microsoft Word](https://github.com/nvaccess/nvda/issues/7409)
- [Issue #11077 - Restrict event processing to objects of interest](https://github.com/nvaccess/nvda/issues/11077)
- [Issue #11209 - UIA selective event registration](https://github.com/nvaccess/nvda/issues/11209)
- [Issue #3578 - PowerPoint 2013 tabbing issues](https://github.com/nvaccess/nvda/issues/3578)
- [Issue #4850 - PowerPoint slideshow reading](https://github.com/nvaccess/nvda/issues/4850)
- [Discussion #13784 - NVDA IAccessible vs UIA](https://github.com/nvaccess/nvda/discussions/13784)

### 11.3 Microsoft Documentation

- [UI Automation and Active Accessibility](https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-msaa) - MSAA/UIA relationship
- [Obtaining UI Automation Elements](https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-obtainingelements) - Element retrieval patterns

### 11.4 Third-Party Resources

- [Python-UIAutomation-for-Windows](https://github.com/yinkaisheng/Python-UIAutomation-for-Windows) - Python UIA wrapper
- [NVDA Add-on Development Guide](https://github.com/nvdaaddons/DevGuide/wiki/NVDA-Add-on-Development-Guide) - Development guide
- [NVDA API Documentation](https://nvda-kr.github.io/nvdaapi/) - Module hierarchy
- [Stack Overflow - IAccessible vs UIA](https://stackoverflow.com/questions/55129774/what-is-the-difference-between-iaccessible-iaccessible2-uiautomation-and-msaa) - API differences

### 11.5 NVDA Add-ons for Development

- [Event Tracker](https://addons.nvda-project.org/addons/evtTracker.en.html) - Monitor events during development
- [NVDA Dev & Test Toolbox](https://addons.nvda-project.org/addons/nvdaDevTestToolbox.en.html) - Development utilities

---

## Summary of Key Recommendations for PowerPoint Plugin

1. **Do NOT use UIA for slide content** - PowerPoint's UIA implementation is incomplete; use COM automation instead

2. **Consider UIA for task panes** - Modern task panes (Accessibility Checker, Design Ideas) may work better with UIA

3. **Handle ribbon carefully** - NetUIHWND windows have focus event issues; test thoroughly

4. **Use hybrid approach** - COM for document content, UIA for modern UI panels

5. **Always handle COMError** - UIA elements can become invalid; wrap all UIA calls in try/except

6. **Cache strategically** - Use cache requests to minimize COM calls for frequently accessed properties

7. **Test with actual screen reader workflow** - Automated tests may not catch all accessibility issues

---

*Document Version: 1.0*
*Research Date: December 2025*
*Author: AI Research Assistant*
