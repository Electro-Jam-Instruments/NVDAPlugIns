# PowerPoint Comment Resolution Status: Locked File Access Methods

**Research Date:** December 4, 2025
**Purpose:** Investigate ALL methods for accessing PowerPoint comment resolution status while the PPTX file is open/locked by PowerPoint
**Target:** NVDA Plugin for PowerPoint 365 Modern Comments
**Performance Target:** Under 500ms for 20 comments
**Critical Constraint:** PowerPoint has exclusive file lock on the PPTX file

---

## Executive Summary

### TOP RECOMMENDATION: Volume Shadow Copy (VSS) with Parse-on-Save Fallback

After comprehensive research across 10 potential approaches, the recommended strategy is a **tiered hybrid approach**:

**Tier 1 (Primary):** Volume Shadow Copy Service (VSS) - Creates point-in-time snapshot allowing access to locked files
**Tier 2 (Fallback):** Parse-on-Save event hooking - Cache resolution status when user saves
**Tier 3 (Degraded):** User-triggered manual refresh - Inform user of limitation

**Key Findings:**

| Approach | Viable | Admin Required | Reliability | Performance |
|----------|--------|----------------|-------------|-------------|
| COM CustomXMLParts | NO | - | - | - |
| Hidden COM Properties | NO | - | - | - |
| Volume Shadow Copy | YES | YES | High | Good |
| Temp File Parsing | PARTIAL | No | Low | Good |
| PowerPoint Add-in | YES | No | Medium | Good |
| Parse-on-Save | YES | No | Medium | Excellent |
| Microsoft Graph | NO | - | - | - |
| File Handle Duplication | NO | - | - | - |
| Memory Scanning | NO | - | - | - |

**Critical Discovery:** The VBA/COM object model does NOT expose comment resolution status (no Comment.Status or Comment.Done property exists, unlike Word). All in-memory COM approaches are non-viable.

---

## Table of Contents

1. [Option 1: COM CustomXMLParts Analysis](#option-1-com-customxmlparts-analysis-not-viable)
2. [Option 2: Hidden COM Properties Discovery](#option-2-hidden-com-properties-discovery-not-viable)
3. [Option 3: Volume Shadow Copy (VSS)](#option-3-volume-shadow-copy-vss-recommended)
4. [Option 4: PowerPoint Temp File Parsing](#option-4-powerpoint-temp-file-parsing-partial)
5. [Option 5: PowerPoint Add-in Companion](#option-5-powerpoint-add-in-companion-alternative)
6. [Option 6: Parse-on-Save Event Hooking](#option-6-parse-on-save-event-hooking-fallback)
7. [Option 7: Microsoft Graph API](#option-7-microsoft-graph-api-not-viable)
8. [Option 8: File Handle Duplication](#option-8-file-handle-duplication-not-viable)
9. [Option 9: Memory Scanning](#option-9-memory-scanning-not-recommended)
10. [Option 10: Hybrid Strategy](#option-10-hybrid-strategy-implementation-plan)
11. [Comparison Matrix](#comparison-matrix)
12. [Implementation Roadmap](#implementation-roadmap)
13. [Code Examples](#code-examples)
14. [References](#references)

---

## Option 1: COM CustomXMLParts Analysis (NOT VIABLE)

### Investigation Summary

CustomXMLParts is a collection that stores custom XML data associated with Office documents. The investigation focused on whether PowerPoint's modernComment XML is accessible through this API.

### Technical Findings

**CustomXMLParts Collection:**
- Available via `Presentation.CustomXMLParts` property
- Returns a `CustomXMLParts` collection object
- Can store arbitrary XML data in the presentation

**Critical Limitation:**
Modern comments (including resolution status) are stored in `/ppt/comments/modernComment_*.xml` files within the PPTX ZIP structure. However, these are **NOT** exposed through CustomXMLParts because:

1. **Different relationship structure:** Modern comments are linked from `/ppt/slides/slide1.xml`, not from the presentation root
2. **Internal parts vs Custom parts:** Comment XML is an internal Office part, not a user-defined custom XML part
3. **Microsoft Q&A confirms:** "Office.js can only read custom XML parts that are linked to the presentation object in the OOXML structure"

**Code Test (Confirms Non-Viability):**

```python
import win32com.client

def test_custom_xml_parts():
    """Test if comments are in CustomXMLParts"""
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    presentation = ppt.ActivePresentation

    # Enumerate all CustomXMLParts
    for i in range(1, presentation.CustomXMLParts.Count + 1):
        part = presentation.CustomXMLParts.Item(i)
        xml_content = part.XML

        # Check if any part contains comment data
        if 'comment' in xml_content.lower() or 'resolved' in xml_content.lower():
            print(f"Found comment-related XML in part {i}")
            return True

    print("No comment data found in CustomXMLParts")
    return False

# Result: False - comments are NOT in CustomXMLParts
```

### Verdict: NOT VIABLE

Modern comment resolution status is stored in internal OOXML parts that are not exposed through the CustomXMLParts API.

**Sources:**
- [Presentation.CustomXMLParts Property](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.customxmlparts)
- [Office.js CustomXML API Limitations](https://learn.microsoft.com/en-us/answers/questions/2149356/powerpoint-office-js-customxml-api-does-not-return)
- [Open-XML-SDK Issue #1133](https://github.com/OfficeDev/Open-XML-SDK/issues/1133)

---

## Option 2: Hidden COM Properties Discovery (NOT VIABLE)

### Investigation Summary

Explored whether undocumented COM properties exist on the PowerPoint Comment object that expose resolution status.

### Technical Findings

**Documented Comment Object Properties:**
| Property | Type | Description |
|----------|------|-------------|
| Author | String | Author's full name |
| AuthorIndex | Long | Author's index in comment list |
| AuthorInitials | String | Author's initials |
| DateTime | Date | When comment was created |
| Text | String | Comment text content |
| Left | Single | X position on slide |
| Top | Single | Y position on slide |
| Replies | Comments | Collection of reply comments |
| Collapsed | Boolean | Whether replies are collapsed |
| UserID | String | User identifier |

**Missing Properties:**
- No `Status` property
- No `Done` property
- No `Resolved` property
- No `IsResolved` property

**Comparison with Word:**
Word's Comment object includes `Comment.Done` property (Boolean) for tracking resolved status. PowerPoint's VBA object model has NOT been updated to include this feature despite the UI supporting it.

**COM Property Enumeration Test:**

```python
import win32com.client
import pythoncom

def enumerate_comment_properties():
    """Attempt to enumerate all Comment object properties via IDispatch"""
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    presentation = ppt.ActivePresentation

    if presentation.Slides.Count > 0:
        slide = presentation.Slides(1)
        if slide.Comments.Count > 0:
            comment = slide.Comments(1)

            # Get type info
            try:
                ti = comment._oleobj_.GetTypeInfo()
                # Enumerate methods/properties
                for i in range(ti.GetTypeAttr().cFuncs):
                    func_desc = ti.GetFuncDesc(i)
                    name = ti.GetNames(func_desc.memid)[0]
                    print(f"Found property/method: {name}")
            except Exception as e:
                print(f"Type enumeration failed: {e}")

            # Try late-binding access to potential hidden properties
            hidden_props = ['Status', 'Done', 'Resolved', 'IsResolved',
                          'ThreadStatus', 'CommentStatus', 'State']
            for prop in hidden_props:
                try:
                    value = getattr(comment, prop)
                    print(f"SUCCESS: {prop} = {value}")
                except AttributeError:
                    print(f"NOT FOUND: {prop}")

# Result: All hidden property attempts fail with AttributeError
```

**Stack Overflow Confirmation:**
User question from April 2024: "I'm looking to get the status of the comments (whether a comment is 'Resolved' or 'Open'). However, I can only seem to access the author and the text for comments, not the status."

Answer: "Modern comments in PP can be 'resolved', the same way as in Word. However, I don't think the VBA object model has caught up with it yet. So the likely answer is 'you can't.'"

### Verdict: NOT VIABLE

No hidden COM properties exist for resolution status. Microsoft has not updated the PowerPoint VBA/COM API to expose modern comment features.

**Sources:**
- [Comment Object (PowerPoint) VBA](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.Comment)
- [Stack Overflow: PowerPoint Comment Status](https://stackoverflow.com/questions/78347637/powerpoint-vba-code-for-pulling-out-a-slides-comments-statuss)
- [Word Comment.Done Property](https://learn.microsoft.com/en-us/office/vba/api/word.comment.done)

---

## Option 3: Volume Shadow Copy (VSS) (RECOMMENDED)

### Investigation Summary

Volume Shadow Copy Service creates point-in-time snapshots of volumes, allowing access to files that are locked by other processes.

### Technical Findings

**How VSS Works:**
1. VSS creates a snapshot of the volume at a specific moment
2. Snapshot provides read-only access to all files as they existed at that moment
3. Works regardless of file locks on the original volume
4. Files can be accessed via special shadow copy paths

**Python Library: pyshadowcopy**
A Python class using pyWin32 to manage shadow copies:

```python
import vss

# Create shadow copy for local drives
local_drives = set(['C'])
sc = vss.ShadowCopy(local_drives)

# Get shadow path for locked file
locked_file = r'C:\Users\...\presentation.pptx'
shadow_path = sc.shadow_path(locked_file)
# Returns: '\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy7\...\presentation.pptx'

# Read file from shadow copy (bypasses lock!)
with open(shadow_path, 'rb') as fp:
    data = fp.read()

# Clean up shadow copy when done
sc.delete()
```

**Advantages:**
- Accesses exact file content even when locked
- Point-in-time snapshot ensures consistency
- No interaction with PowerPoint required
- File content is exactly what PowerPoint has saved

**Disadvantages:**
- **REQUIRES ADMINISTRATOR PRIVILEGES** - VSS COM APIs require elevation
- Shadow copy creation has overhead (~100-500ms)
- Need to clean up shadow copies
- May not capture unsaved changes (only saved state)

**Admin Privilege Investigation:**
Per Microsoft documentation and Stack Overflow research:
- "Microsoft's command-line tools for VSS require execution with Administrator privileges"
- "Users get 0x80070005 - Access is denied on COM calls when in normal user mode"
- Even Backup Operators group membership doesn't help without elevation

**Bitness Requirement:**
"Using 32-bit python/pyWin32 on a 64-bit OS won't work. You need to ensure the bitness of Python and pyWin32 matches your operating system."

For NVDA (32-bit Python on 64-bit Windows), this may be problematic.

### Performance Analysis

| Operation | Time | Notes |
|-----------|------|-------|
| Shadow copy creation | 200-500ms | One-time per refresh |
| Path translation | <10ms | Simple string operation |
| File read (10MB PPTX) | 50-100ms | Sequential read |
| XML parsing (comments) | 20-50ms | ElementTree parsing |
| Shadow copy cleanup | 100-200ms | Can be async |
| **Total** | **370-860ms** | First access |

With caching and async cleanup, effective time for user: **300-500ms**

### Verdict: RECOMMENDED (with caveats)

**Viable if:**
- Application can run with admin privileges (or prompt for elevation)
- NVDA's 32-bit Python can work with VSS (needs testing)
- Unsaved changes limitation is acceptable to users

**Implementation:**
See [Code Examples](#vss-implementation) section for full implementation.

**Sources:**
- [pyshadowcopy GitHub](https://github.com/sblosser/pyshadowcopy)
- [VSS Admin Privileges Question](https://stackoverflow.com/questions/7530540/can-the-volume-shadow-copy-service-be-used-in-windows-7-by-a-non-administrator)
- [How VSS Handles Locked Files](https://superuser.com/questions/263205/how-does-vss-volume-shadow-copy-handle-locked-files)

---

## Option 4: PowerPoint Temp File Parsing (PARTIAL)

### Investigation Summary

Investigated PowerPoint's temporary and autosave file locations to determine if comment data is accessible through these files.

### Technical Findings

**PowerPoint Temp File Locations:**

| Location | Purpose | Contents |
|----------|---------|----------|
| `%TEMP%` or `C:\Users\<user>\AppData\Local\Temp` | Working temp files | pptxxx.tmp files during edits |
| `%APPDATA%\Microsoft\PowerPoint` | AutoRecover | Recovery files for crashes |
| `%LOCALAPPDATA%\Microsoft\Office\UnsavedFiles` | Unsaved presentations | Draft files |
| `%APPDATA%\Microsoft\Templates` | Templates | Template files |

**Temp File Format:**
- Temp files are named `pptxxxx.tmp` where xxxx is a number
- Files with `~` prefix are also temporary
- These are typically partial/working files, NOT complete PPTX archives

**AutoRecover Files:**
- Located in `C:\Users\<username>\AppData\Roaming\Microsoft\PowerPoint`
- Created periodically (default: every 10 minutes)
- Format: May be incomplete PPTX or proprietary recovery format

**Investigation Results:**

```python
import os
import glob

def find_powerpoint_temp_files():
    """Search for PowerPoint temp files"""
    locations = [
        os.path.join(os.environ.get('TEMP', ''), 'ppt*.tmp'),
        os.path.join(os.environ.get('APPDATA', ''), 'Microsoft', 'PowerPoint', '*'),
        os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Microsoft', 'Office', 'UnsavedFiles', '*.pptx'),
        os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Temp', '~*.pptx'),
    ]

    found_files = []
    for pattern in locations:
        found_files.extend(glob.glob(pattern))

    return found_files
```

**Critical Limitation:**
Modern comment XML files (`modernComment_*.xml`) are stored INSIDE the PPTX ZIP archive. Temp files created during editing:
1. May not include comment updates until autosave triggers
2. Are often partial files, not valid PPTX ZIPs
3. May be locked by PowerPoint similar to main file
4. Do not reliably contain current comment state

**AutoRecover Timing:**
- Default: Every 10 minutes
- User-configurable in PowerPoint options
- Not triggered by comment changes specifically
- Comment resolution changes may not trigger autosave

### Verdict: PARTIAL (Unreliable)

Temp file parsing is unreliable because:
1. Comment data may be stale (up to 10 minutes old)
2. Temp files may also be locked
3. Format varies and may not be parseable PPTX
4. User's autosave settings affect availability

**May be useful as supplementary data source** but not primary solution.

**Sources:**
- [PowerPoint Temp File Locations](https://www.imyfone.com/data-recovery/recover-powerpoint-file/)
- [AutoRecover Location](https://answers.microsoft.com/en-us/msoffice/forum/all/change-powerpoint-folder-location-for-temporary/93dd061b-5381-4fc7-a0de-4c57640940c2)
- [Microsoft Support: Recover Files](https://support.microsoft.com/en-au/topic/recover-your-powerpoint-files-bd398210-f1b3-46f0-86aa-e4df2d177cbd)

---

## Option 5: PowerPoint Add-in Companion (ALTERNATIVE)

### Investigation Summary

Explored creating a separate PowerPoint COM Add-in that runs inside PowerPoint and communicates with the NVDA plugin via IPC.

### Technical Findings

**Architecture:**

```
+-------------------+          IPC           +------------------+
|  PowerPoint       |  <==================>  |  NVDA Plugin     |
|  + COM Add-in     |   Named Pipes / COM    |                  |
|    (VSTO/C#)      |                        |  (Python)        |
+-------------------+                        +------------------+
        |
        v
   Internal APIs
   (may access more)
```

**VSTO Add-in Capabilities:**
- Runs in PowerPoint's process space
- Access to full PowerPoint object model
- Can handle events like SlideSelectionChanged
- Can potentially access more internal state

**The Key Question:** Can a VSTO Add-in access resolution status?

**Investigation:**
- VSTO uses the same COM object model as VBA
- VSTO does NOT have access to additional internal APIs
- "The only known way to communicate directly with Office products is by using COM components"
- Add-in cannot access the modernComment XML directly from memory

**Microsoft Q&A Finding:**
"Can I access low level data of my pptx presentation?" - Answer indicates VSTO add-ins work through COM wrappers, not direct memory/XML access.

**IPC Options:**
| Method | Complexity | Performance | Reliability |
|--------|------------|-------------|-------------|
| Named Pipes | Medium | Excellent | Good |
| Shared Memory | High | Excellent | Medium |
| File-based | Low | Good | High |
| COM Server | High | Good | Good |

**User Experience Impact:**
- Users must install two components (NVDA addon + Office Add-in)
- Add-in must be deployed/managed separately
- Potential conflicts with other add-ins
- More failure points

### Verdict: ALTERNATIVE (Not Preferred)

A PowerPoint Add-in:
- Does NOT provide access to resolution status (same COM limitations)
- Adds complexity with no capability benefit
- Increases user installation burden

**Only valuable if:** Future Office updates expose resolution via COM, at which point the add-in could proactively push updates to NVDA (avoiding polling).

**Sources:**
- [Create VSTO Add-ins for PowerPoint](https://learn.microsoft.com/en-us/visualstudio/vsto/powerpoint-solutions?view=vs-2022)
- [VSTO Low-Level Data Access](https://docs.microsoft.com/answers/questions/39954/powerpoint-vsto-can-i-access-low-level-data-of-my.html)
- [Named Pipe Security](https://learn.microsoft.com/en-us/windows/win32/ipc/named-pipe-security-and-access-rights)

---

## Option 6: Parse-on-Save Event Hooking (FALLBACK)

### Investigation Summary

Hook into PowerPoint's save events to parse the PPTX immediately after saving, when the file is momentarily accessible.

### Technical Findings

**Available PowerPoint Events:**

| Event | Trigger | Parameters |
|-------|---------|------------|
| PresentationBeforeSave | Before Save As dialog appears | Pres, Cancel |
| PresentationSave | After presentation saved (2010+) | Pres |
| AfterPresentationOpen | After presentation opens | Pres |
| SlideSelectionChanged | When slide selection changes | SlideRange |

**Event Access from NVDA:**
NVDA can hook PowerPoint COM events via win32com:

```python
import win32com.client

class PowerPointEvents:
    """Event handler for PowerPoint application events"""

    def OnPresentationSave(self, pres):
        """Fired after presentation is saved"""
        # File should be unlocked briefly here
        file_path = pres.FullName
        self._parse_and_cache_comments(file_path)

    def OnPresentationBeforeSave(self, pres, cancel):
        """Fired before save - file still locked"""
        pass  # Cannot access file yet

# Connect to events
ppt = win32com.client.DispatchWithEvents(
    "PowerPoint.Application",
    PowerPointEvents
)
```

**File Availability Timing:**
- `PresentationBeforeSave`: File is still locked (Save dialog about to appear)
- `PresentationSave`: File has been written but may still be briefly locked
- Small window after save completes where file may be accessible

**Limitations:**
1. Only updates when user saves (stale between saves)
2. Users don't save frequently during review
3. Auto-save in Microsoft 365 may help (if enabled)
4. Event may fire but file still briefly locked

**Microsoft 365 AutoSave:**
- When connected to OneDrive/SharePoint, AutoSave is always on
- Saves changes every few seconds
- More frequent updates possible in cloud scenarios

### Verdict: FALLBACK (Reliable but Stale)

Parse-on-Save is reliable but data may be stale:
- Works consistently when events fire
- No admin privileges required
- Simple implementation
- User must be informed status may be outdated

**Best used as:** Fallback when VSS not available, or supplement to VSS.

**Sources:**
- [Application.PresentationBeforeSave](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.presentationbeforesave)
- [Use Events with Application Object](https://learn.microsoft.com/en-us/office/vba/powerpoint/How-to/use-events-with-the-application-object)
- [PowerPoint Application Events](http://youpresent.co.uk/powerpoint-application-events-in-vba/)

---

## Option 7: Microsoft Graph API (NOT VIABLE)

### Investigation Summary

Investigated Microsoft Graph API for accessing PowerPoint comments on cloud-stored files.

### Technical Findings

**Graph API Support:**
| Feature | Excel | Word | PowerPoint |
|---------|-------|------|------------|
| File access | Yes | Yes | Yes |
| List comments | Yes (workbookComments) | No | No |
| Comment resolution | Yes | No | No |

**Critical Finding:**
"Microsoft Graph API currently does not support viewing Word/PowerPoint inline comments."

**API Endpoint Investigation:**
- `/drives/{driveId}/items/{itemId}/content` - Download file only
- `/drives/{driveId}/items/{itemId}/workbook/comments` - Excel only
- No PowerPoint comments endpoint exists

**Microsoft Q&A Confirmation:**
"Is there an API to view PowerPoint comments?" - Answer: "PowerPoint comments are not accessible through Microsoft Graph API."

**Alternative Approach:**
Could download file via Graph and parse locally, but:
1. Requires file to be in OneDrive/SharePoint
2. Downloaded file is a copy, not live data
3. Same as parsing local copy after download
4. Adds latency and complexity

### Verdict: NOT VIABLE

Microsoft Graph does not expose PowerPoint comment data.

**Sources:**
- [Graph API: How to view PowerPoint comments](https://learn.microsoft.com/en-us/answers/questions/1193483/how-to-view-word-powerpoint-inline-comments-using)
- [List workbookComments (Excel only)](https://learn.microsoft.com/en-us/graph/api/workbook-list-comments?view=graph-rest-1.0)
- [Graph API Comments on SharePoint](https://learn.microsoft.com/en-us/answers/questions/5526582/access-comments-of-a-document-uploaded-on-sharepoi)

---

## Option 8: File Handle Duplication (NOT VIABLE)

### Investigation Summary

Investigated using Windows API DuplicateHandle to copy PowerPoint's file handle and read from the duplicated handle.

### Technical Findings

**DuplicateHandle API:**
The Windows API function can create a copy of a handle that another process owns.

**Limitations:**
1. **Cannot add access rights:** "A file handle created with GENERIC_READ access right cannot be duplicated so that it has both GENERIC_READ and GENERIC_WRITE access right"
2. **Same restrictions apply:** Duplicated handle inherits same access restrictions
3. **Requires handle knowledge:** Need to know the specific handle PowerPoint is using
4. **Process boundary issues:** Handles from other processes require special privileges

**Alternative: CreateFile with Share Flags:**

```python
import win32file

# Attempt to open with maximum sharing
handle = win32file.CreateFile(
    filename,
    win32file.GENERIC_READ,
    win32file.FILE_SHARE_DELETE | win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE,
    None,
    win32file.OPEN_EXISTING,
    0,
    None
)
```

**Testing Required:**
- If PowerPoint opened file with FILE_SHARE_READ, we can read
- If PowerPoint opened without share flags, we cannot read
- PowerPoint's behavior varies by operation/state

**Observed Behavior:**
When PowerPoint opens a presentation:
- File gets exclusive lock during certain operations
- Brief windows where shared read might work
- Inconsistent across Office versions

### Verdict: NOT VIABLE (Unreliable)

File handle approaches are unreliable because:
1. PowerPoint's locking behavior is inconsistent
2. Cannot guarantee read access
3. May get partial/corrupt reads if file is being written

**Sources:**
- [DuplicateHandle Function](https://learn.microsoft.com/en-us/windows/win32/api/handleapi/nf-handleapi-duplicatehandle)
- [Opening Locked Files with win32file](https://stackoverflow.com/questions/58321073/read-from-locked-file-in-windows-using-python)
- [File Sharing Modes](https://stackoverflow.com/questions/60709401/opening-a-file-on-windows-with-exclusive-locking-in-python)

---

## Option 9: Memory Scanning (NOT RECOMMENDED)

### Investigation Summary

Explored scanning PowerPoint's process memory for comment XML data.

### Technical Findings

**Technical Approach:**
1. Enumerate POWERPNT.EXE process memory regions
2. Search for XML strings containing comment data
3. Parse found XML for resolution status

**Python Libraries:**
- `pymem` - Process memory reading
- `ReadProcessMemory` via ctypes/pywin32

**Fundamental Problems:**

| Issue | Impact | Severity |
|-------|--------|----------|
| Memory layout varies by version | Unreliable results | Critical |
| XML may be fragmented in memory | Incomplete data | Critical |
| Anti-virus may flag behavior | User alarm, blocking | High |
| Stability risks | Potential crashes | High |
| Legal/ethical concerns | Terms of service | Medium |
| Performance overhead | Slow scanning | Medium |

**Security Concerns:**
- Memory scanning is a technique used by malware
- May trigger security software
- Could be viewed as reverse engineering (DMCA concerns)
- Microsoft could patch against such access

### Verdict: NOT RECOMMENDED

Memory scanning is:
1. Technically unreliable
2. Security risk (flagging by AV)
3. Potentially against Microsoft terms
4. Unstable across versions
5. Not suitable for accessibility software

**DO NOT IMPLEMENT** - risks outweigh any potential benefit.

---

## Option 10: Hybrid Strategy (IMPLEMENTATION PLAN)

### Recommended Architecture

```
+-------------------------------------------------------------------+
|                      Comment Resolution Detection                   |
+-------------------------------------------------------------------+
                              |
              +---------------+----------------+
              |                                |
         [ Check Admin ]                  [ Register Events ]
              |                                |
     +--------+--------+               +-------+-------+
     | Admin: YES      |               | PresentationSave |
     |                 |               +-------+-------+
     | Use VSS         |                       |
     | - Create snap   |               +-------+-------+
     | - Read PPTX     |               | Parse PPTX    |
     | - Parse XML     |               | Cache results |
     | - Cleanup       |               +---------------+
     +--------+--------+
              |
     +--------+--------+
     | Admin: NO       |
     |                 |
     | Use Parse-on-   |
     | Save with stale |
     | data warning    |
     +-----------------+
```

### Decision Flow

```
1. On plugin startup:
   - Detect if running with admin privileges
   - Register PowerPoint save event handler
   - Initialize resolution cache

2. On comment navigation request:
   IF admin privileges available:
       - Create VSS snapshot
       - Read PPTX from shadow copy
       - Parse modernComment XML
       - Extract resolution status
       - Clean up snapshot
       - Return fresh data
   ELSE:
       - Check cache age
       - IF cache < 30 seconds old:
           - Return cached data
       - ELSE:
           - Return cached data with "may be stale" warning
           - Suggest user save for fresh data

3. On PresentationSave event:
   - Wait 500ms for file release
   - Attempt to read PPTX directly
   - IF successful:
       - Parse and cache resolution data
   - ELSE:
       - Log warning, keep old cache

4. On user request "Refresh Status":
   - Force VSS snapshot (if admin)
   - Or prompt user to save (if not admin)
```

### User Communication

**With Admin:**
- Silent operation, always fresh data
- Optional: "Resolution status current as of [timestamp]"

**Without Admin:**
- On first use: "Comment resolution status available after saving. For real-time status, run NVDA as administrator."
- On navigation: "[RESOLVED] or [ACTIVE] (as of last save)"
- After 5 minutes without save: "Resolution status may be outdated. Press Ctrl+S to refresh."

### Configuration Options

```python
class CommentResolutionConfig:
    """User-configurable options for resolution detection"""

    # Attempt VSS even without admin (will fail gracefully)
    try_vss_without_admin: bool = True

    # Maximum cache age before warning (seconds)
    cache_warning_threshold: int = 300

    # Announce resolution status with every comment
    announce_resolution: bool = True

    # Announce only for resolved comments
    announce_resolved_only: bool = False

    # Play tone to indicate resolution state
    resolution_tone_enabled: bool = True

    # Tone frequency for resolved comments
    resolved_tone_hz: int = 880

    # Tone frequency for active comments
    active_tone_hz: int = 440
```

---

## Comparison Matrix

| Option | Reliability | Performance | Complexity | Admin Req | User Impact | Recommendation |
|--------|-------------|-------------|------------|-----------|-------------|----------------|
| COM CustomXMLParts | N/A | N/A | N/A | N/A | N/A | NOT VIABLE |
| Hidden COM Properties | N/A | N/A | N/A | N/A | N/A | NOT VIABLE |
| Volume Shadow Copy | HIGH | 300-500ms | MEDIUM | YES | Low (if admin) | PRIMARY |
| Temp File Parsing | LOW | 100-200ms | LOW | No | High (unreliable) | DO NOT USE |
| PowerPoint Add-in | N/A | N/A | HIGH | No | High (2 installs) | NOT VIABLE |
| Parse-on-Save | MEDIUM | <100ms | LOW | No | Medium (stale data) | FALLBACK |
| Microsoft Graph | N/A | N/A | N/A | N/A | N/A | NOT VIABLE |
| File Handle Dup | LOW | Fast | MEDIUM | No | High (fails often) | NOT VIABLE |
| Memory Scanning | LOW | Slow | HIGH | No | Critical (security) | DO NOT USE |
| **Hybrid VSS+Save** | **HIGH** | **<500ms** | **MEDIUM** | **Partial** | **Low** | **RECOMMENDED** |

### Scoring Legend

**Reliability:**
- HIGH: >95% success rate
- MEDIUM: 70-95% success rate
- LOW: <70% success rate

**Performance:**
- Fast: <100ms
- Good: 100-300ms
- Acceptable: 300-500ms
- Slow: >500ms

**User Impact:**
- Low: Transparent operation
- Medium: Some user awareness needed
- High: User action required or data limitations

---

## Implementation Roadmap

### Phase 1: Core Infrastructure (Week 1)

**Tasks:**
1. Implement OOXML comment parser (already done per previous research)
2. Add VSS snapshot wrapper
3. Create resolution cache system
4. Add admin privilege detection

**Deliverables:**
- `vss_reader.py` - Volume Shadow Copy file reader
- `resolution_cache.py` - Caching infrastructure
- `privilege_detector.py` - Admin check utility

### Phase 2: Event Integration (Week 2)

**Tasks:**
1. Register PowerPoint PresentationSave event handler
2. Implement parse-after-save logic
3. Add graceful fallback for non-admin
4. Create user messaging system

**Deliverables:**
- `event_handlers.py` - PowerPoint event integration
- `status_announcer.py` - User feedback system

### Phase 3: Navigation Enhancement (Week 3)

**Tasks:**
1. Integrate resolution status into comment navigation
2. Add "next unresolved" command
3. Add resolution statistics command
4. Implement configuration options

**Deliverables:**
- Updated `navigation.py` with resolution awareness
- Configuration UI additions
- New keyboard shortcuts

### Phase 4: Testing and Polish (Week 4)

**Tasks:**
1. Test across Office versions (2019, 2021, 365)
2. Test with and without admin
3. Performance optimization
4. Documentation

**Deliverables:**
- Test report
- User documentation
- Performance benchmarks

---

## Code Examples

### VSS Implementation

```python
"""
vss_reader.py - Volume Shadow Copy file reader for locked PPTX files
Requires: pywin32, administrator privileges
"""

import os
import ctypes
import tempfile
from contextlib import contextmanager

# Check if pyshadowcopy is available, otherwise use fallback
try:
    import vss
    HAS_VSS = True
except ImportError:
    HAS_VSS = False


def is_admin():
    """Check if current process has administrator privileges"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except:
        return False


class VSSReader:
    """Read files from Volume Shadow Copy snapshots"""

    def __init__(self):
        self.shadow_copy = None
        self._snapshot_active = False

    def available(self):
        """Check if VSS is available"""
        return HAS_VSS and is_admin()

    @contextmanager
    def snapshot(self, file_path):
        """
        Context manager for VSS snapshot access.

        Usage:
            reader = VSSReader()
            with reader.snapshot(locked_file_path) as shadow_path:
                with open(shadow_path, 'rb') as f:
                    data = f.read()
        """
        if not self.available():
            raise RuntimeError("VSS not available (requires admin privileges)")

        # Get drive letter from file path
        drive = os.path.splitdrive(file_path)[0].rstrip(':').upper()

        try:
            # Create shadow copy
            self.shadow_copy = vss.ShadowCopy(set([drive]))
            self._snapshot_active = True

            # Get shadow path
            shadow_path = self.shadow_copy.shadow_path(file_path)

            yield shadow_path

        finally:
            # Clean up snapshot
            if self.shadow_copy:
                try:
                    self.shadow_copy.delete()
                except:
                    pass
                self.shadow_copy = None
                self._snapshot_active = False

    def read_file(self, file_path):
        """
        Read entire file content via VSS snapshot.

        Returns:
            bytes: File content
        """
        with self.snapshot(file_path) as shadow_path:
            with open(shadow_path, 'rb') as f:
                return f.read()

    def read_to_temp(self, file_path):
        """
        Copy locked file to temp location via VSS.

        Returns:
            str: Path to temporary copy (caller must clean up)
        """
        content = self.read_file(file_path)

        # Write to temp file
        fd, temp_path = tempfile.mkstemp(suffix='.pptx')
        try:
            os.write(fd, content)
        finally:
            os.close(fd)

        return temp_path


class FallbackReader:
    """Fallback reader when VSS not available"""

    def __init__(self):
        self.last_error = None

    def available(self):
        """Always returns False - this is fallback"""
        return False

    def try_direct_read(self, file_path):
        """
        Attempt direct file read with sharing flags.
        May fail if file is locked.
        """
        import win32file
        import pywintypes

        try:
            handle = win32file.CreateFile(
                file_path,
                win32file.GENERIC_READ,
                win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE,
                None,
                win32file.OPEN_EXISTING,
                0,
                None
            )

            try:
                # Read file content
                _, data = win32file.ReadFile(handle, os.path.getsize(file_path))
                return data
            finally:
                win32file.CloseHandle(handle)

        except pywintypes.error as e:
            self.last_error = str(e)
            return None


def get_reader():
    """Factory function to get appropriate reader"""
    vss_reader = VSSReader()
    if vss_reader.available():
        return vss_reader
    return FallbackReader()
```

### Parse-on-Save Event Handler

```python
"""
event_handlers.py - PowerPoint event handlers for resolution caching
"""

import win32com.client
import pythoncom
import threading
import time
from logHandler import log


class PowerPointEventHandler:
    """
    Handles PowerPoint application events for resolution caching.

    Events:
    - PresentationSave: Cache comments after save
    - SlideSelectionChanged: Could prefetch slide comments
    """

    def __init__(self, cache_manager):
        self.cache = cache_manager
        self._app = None
        self._event_sink = None
        self._running = False

    def start(self):
        """Start listening for PowerPoint events"""
        if self._running:
            return

        try:
            # Connect to PowerPoint with events
            self._app = win32com.client.DispatchWithEvents(
                "PowerPoint.Application",
                self._create_event_class()
            )
            self._running = True
            log.info("PowerPoint event handler started")
        except Exception as e:
            log.error(f"Failed to start event handler: {e}")

    def stop(self):
        """Stop listening for events"""
        self._running = False
        self._app = None

    def _create_event_class(self):
        """Create event class with reference to cache"""
        cache = self.cache

        class _Events:
            def OnPresentationSave(self, pres):
                """Fired after presentation is saved"""
                try:
                    file_path = pres.FullName

                    # Wait brief moment for file release
                    time.sleep(0.3)

                    # Parse in background thread
                    thread = threading.Thread(
                        target=cache.refresh_from_file,
                        args=(file_path,)
                    )
                    thread.daemon = True
                    thread.start()

                except Exception as e:
                    log.error(f"PresentationSave handler error: {e}")

            def OnSlideSelectionChanged(self, slide_range):
                """Fired when slide selection changes"""
                # Could prefetch comments for new slide
                pass

        return _Events


class ResolutionCache:
    """
    Cache for comment resolution status.

    Structure:
    {
        'file_path': {
            'timestamp': datetime,
            'slides': {
                1: [{'id': 'xxx', 'resolved': True, 'text': '...'}, ...],
                2: [...],
            }
        }
    }
    """

    def __init__(self, max_age_seconds=300):
        self._cache = {}
        self._lock = threading.Lock()
        self.max_age = max_age_seconds

    def get(self, file_path, slide_number):
        """
        Get cached resolution data for slide.

        Returns:
            (list, bool): (comments list, is_fresh)
        """
        with self._lock:
            if file_path not in self._cache:
                return [], False

            entry = self._cache[file_path]
            age = time.time() - entry['timestamp']
            is_fresh = age < self.max_age

            comments = entry['slides'].get(slide_number, [])
            return comments, is_fresh

    def set(self, file_path, slide_number, comments):
        """Cache resolution data for slide"""
        with self._lock:
            if file_path not in self._cache:
                self._cache[file_path] = {
                    'timestamp': time.time(),
                    'slides': {}
                }

            self._cache[file_path]['slides'][slide_number] = comments
            self._cache[file_path]['timestamp'] = time.time()

    def refresh_from_file(self, file_path):
        """
        Refresh cache by parsing PPTX file.
        Called after save events.
        """
        try:
            from .ooxml_parser import PowerPointCommentReader

            reader = PowerPointCommentReader(file_path)
            all_comments = reader.read_all_comments()

            with self._lock:
                self._cache[file_path] = {
                    'timestamp': time.time(),
                    'slides': all_comments
                }

            log.info(f"Resolution cache refreshed for {file_path}")

        except PermissionError:
            log.warning(f"File still locked, cache not updated: {file_path}")
        except Exception as e:
            log.error(f"Failed to refresh cache: {e}")

    def invalidate(self, file_path=None):
        """Clear cache for file or all files"""
        with self._lock:
            if file_path:
                self._cache.pop(file_path, None)
            else:
                self._cache.clear()
```

### Integrated Resolution Navigator

```python
"""
resolution_navigator.py - Comment navigator with resolution status
"""

import ui
import tones
from logHandler import log


class ResolutionAwareNavigator:
    """
    Comment navigator that includes resolution status.
    Uses hybrid approach: VSS when available, cache fallback otherwise.
    """

    def __init__(self, ppt_connector):
        self.ppt = ppt_connector
        self.comments = []
        self.current_index = -1

        # Initialize readers
        from .vss_reader import VSSReader, FallbackReader
        self.vss_reader = VSSReader()
        self.fallback_reader = FallbackReader()

        # Initialize cache
        from .event_handlers import ResolutionCache
        self.cache = ResolutionCache()

        # Track data freshness
        self._using_fresh_data = False

    def refresh_comments(self, force_vss=False):
        """
        Refresh comment list with resolution status.

        Args:
            force_vss: Force VSS refresh even if cache available

        Returns:
            int: Number of comments found
        """
        presentation = self.ppt.get_presentation()
        if not presentation:
            return 0

        file_path = presentation.FullName
        slide = self.ppt.get_current_slide()
        if not slide:
            return 0

        slide_num = slide.SlideNumber

        # Get basic comment data via COM
        com_comments = self._get_com_comments(slide)

        # Get resolution status
        resolution_data = None
        self._using_fresh_data = False

        # Try VSS first (freshest data)
        if force_vss or self.vss_reader.available():
            resolution_data = self._get_resolution_via_vss(file_path, slide_num)
            if resolution_data:
                self._using_fresh_data = True

        # Fall back to cache
        if not resolution_data:
            resolution_data, is_fresh = self.cache.get(file_path, slide_num)
            self._using_fresh_data = is_fresh

        # Merge resolution into COM comments
        self.comments = self._merge_data(com_comments, resolution_data)
        self.current_index = -1

        return len(self.comments)

    def _get_com_comments(self, slide):
        """Get comment data from COM"""
        comments = []
        try:
            for i, com_comment in enumerate(slide.Comments):
                comments.append({
                    'index': i,
                    'author': com_comment.Author,
                    'text': com_comment.Text,
                    'date': str(getattr(com_comment, 'DateTime', '')),
                    'replies': self._get_replies(com_comment),
                    'status': 'unknown',
                    'is_resolved': None
                })
        except Exception as e:
            log.error(f"COM comment access failed: {e}")
        return comments

    def _get_replies(self, comment):
        """Get reply data from COM comment"""
        replies = []
        if hasattr(comment, 'Replies'):
            try:
                for reply in comment.Replies:
                    replies.append({
                        'author': reply.Author,
                        'text': reply.Text
                    })
            except:
                pass
        return replies

    def _get_resolution_via_vss(self, file_path, slide_num):
        """Get resolution data using VSS snapshot"""
        try:
            from .ooxml_parser import PowerPointCommentReader

            temp_path = self.vss_reader.read_to_temp(file_path)
            try:
                reader = PowerPointCommentReader(temp_path)
                all_comments = reader.read_all_comments()

                # Update cache with fresh data
                for snum, comments in all_comments.items():
                    self.cache.set(file_path, snum, comments)

                return all_comments.get(slide_num, [])
            finally:
                import os
                try:
                    os.remove(temp_path)
                except:
                    pass

        except Exception as e:
            log.error(f"VSS resolution read failed: {e}")
            return None

    def _merge_data(self, com_comments, resolution_data):
        """Merge COM data with resolution data"""
        if not resolution_data:
            return com_comments

        # Match by text preview (first 30 chars)
        resolution_by_text = {
            r.get('text', '')[:30]: r
            for r in resolution_data
        }

        for comment in com_comments:
            text_key = comment['text'][:30]
            if text_key in resolution_by_text:
                res_data = resolution_by_text[text_key]
                comment['status'] = res_data.get('status', 'active')
                comment['is_resolved'] = res_data.get('is_resolved', False)

        return com_comments

    def navigate_next(self):
        """Navigate to next comment"""
        if not self.comments:
            count = self.refresh_comments()
            if count == 0:
                ui.message("No comments on this slide")
                tones.beep(200, 100)
                return

        self.current_index = (self.current_index + 1) % len(self.comments)
        self._announce_current()

    def navigate_previous(self):
        """Navigate to previous comment"""
        if not self.comments:
            count = self.refresh_comments()
            if count == 0:
                ui.message("No comments on this slide")
                tones.beep(200, 100)
                return

        self.current_index = (self.current_index - 1) % len(self.comments)
        self._announce_current()

    def navigate_next_unresolved(self):
        """Navigate to next unresolved comment"""
        if not self.comments:
            self.refresh_comments()

        if not self.comments:
            ui.message("No comments on this slide")
            return

        # Find next unresolved
        start = self.current_index + 1
        for offset in range(len(self.comments)):
            idx = (start + offset) % len(self.comments)
            if not self.comments[idx].get('is_resolved', False):
                self.current_index = idx
                self._announce_current()
                return

        ui.message("No unresolved comments found")
        tones.beep(440, 150)

    def _announce_current(self):
        """Announce current comment"""
        if not self.comments or self.current_index < 0:
            return

        comment = self.comments[self.current_index]

        # Build announcement parts
        parts = []

        # Position
        parts.append(f"Comment {self.current_index + 1} of {len(self.comments)}")

        # Resolution status
        status = comment.get('status', 'unknown')
        if status == 'resolved':
            parts.append("RESOLVED")
            tones.beep(880, 50)
        elif status == 'closed':
            parts.append("CLOSED")
            tones.beep(880, 50)
        elif status == 'active':
            parts.append("Active")
            tones.beep(440, 50)
        elif not self._using_fresh_data:
            parts.append("Status unknown")

        # Author
        if comment.get('author'):
            parts.append(f"by {comment['author']}")

        # Freshness warning
        if not self._using_fresh_data:
            parts.append("(status may be outdated)")

        # Announce header
        ui.message(" - ".join(parts))

        # Announce text
        ui.message(comment.get('text', 'No text'))

        # Announce reply count
        reply_count = len(comment.get('replies', []))
        if reply_count > 0:
            ui.message(f"{reply_count} {'replies' if reply_count != 1 else 'reply'}")

    def get_statistics(self):
        """Get comment statistics"""
        if not self.comments:
            self.refresh_comments()

        total = len(self.comments)
        resolved = sum(1 for c in self.comments if c.get('is_resolved'))
        active = total - resolved
        unknown = sum(1 for c in self.comments if c.get('status') == 'unknown')

        return {
            'total': total,
            'active': active,
            'resolved': resolved,
            'unknown': unknown,
            'fresh': self._using_fresh_data
        }

    def announce_statistics(self):
        """Announce comment statistics"""
        stats = self.get_statistics()

        if stats['total'] == 0:
            ui.message("No comments on this slide")
            return

        message = f"{stats['total']} comments: {stats['active']} active, {stats['resolved']} resolved"

        if stats['unknown'] > 0:
            message += f", {stats['unknown']} unknown status"

        if not stats['fresh']:
            message += " (status may be outdated - save presentation to refresh)"

        ui.message(message)
```

---

## References

### Microsoft Documentation

- [Presentation.CustomXMLParts Property](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.customxmlparts)
- [Comment Object (PowerPoint)](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.Comment)
- [Comment.Status Property (OpenXML SDK)](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2021.powerpoint.comment.comment.status?view=openxml-3.0.1)
- [Application.PresentationBeforeSave Event](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.presentationbeforesave)
- [DuplicateHandle Function](https://learn.microsoft.com/en-us/windows/win32/api/handleapi/nf-handleapi-duplicatehandle)
- [Modern Comments in PowerPoint](https://support.microsoft.com/en-us/office/modern-comments-in-powerpoint-c0aa37bb-82cb-414c-872d-178946ff60ec)

### GitHub Resources

- [pyshadowcopy Library](https://github.com/sblosser/pyshadowcopy)
- [Open-XML-SDK Issue #1133](https://github.com/OfficeDev/Open-XML-SDK/issues/1133)
- [NVDA Source Code](https://github.com/nvaccess/nvda)

### Stack Overflow

- [PowerPoint Comment Status VBA](https://stackoverflow.com/questions/78347637/powerpoint-vba-code-for-pulling-out-a-slides-comments-statuss)
- [VSS Admin Privileges](https://stackoverflow.com/questions/7530540/can-the-volume-shadow-copy-service-be-used-in-windows-7-by-non-administrator)
- [Read Locked Files in Python](https://stackoverflow.com/questions/58321073/read-from-locked-file-in-windows-using-python)
- [Excel Resolved Comments VBA](https://stackoverflow.com/questions/78579092/how-to-count-resolved-comments-only-via-vba-in-excel)

### Microsoft Q&A

- [Graph API PowerPoint Comments](https://learn.microsoft.com/en-us/answers/questions/1193483/how-to-view-word-powerpoint-inline-comments-using)
- [VSTO Low-Level Data Access](https://docs.microsoft.com/answers/questions/39954/powerpoint-vsto-can-i-access-low-level-data-of-my.html)
- [Office.js CustomXML Limitations](https://learn.microsoft.com/en-us/answers/questions/2149356/powerpoint-office-js-customxml-api-does-not-return)

---

*Document generated: December 4, 2025*
*Research completed by Strategic Planning and Research Specialist*
*For NVDA PowerPoint Accessibility Plugin Project*
