# NVDA Plugin Development - Decisions

## Decision Log

### 1. Use comtypes, not pywin32

**Decision:** Use `comtypes` for all COM automation
**Date:** December 2025
**Status:** Final

**Rationale:**
- NVDA uses comtypes internally
- pywin32 has DLL conflicts when loaded in NVDA process
- comtypes is already available in NVDA runtime

**Research:** `research/NVDA_PowerPoint_Native_Support_Analysis.md`

---

### 2. App Module Approach (not Global Plugin alone)

**Decision:** Create `appModules/powerpnt.py` as primary entry point
**Date:** December 2025
**Status:** Final

**Rationale:**
- Integrates with NVDA's existing PowerPoint support
- Can use overlay classes for PowerPoint-specific objects
- Inherits existing COM event infrastructure

**Alternatives Considered:**
- Global Plugin only: Would miss PowerPoint-specific events
- Hybrid: More complex, not needed for MVP

**Research:** `research/NVDA_PowerPoint_Native_Support_Analysis.md` Section 6

---

### 3. NVDA Version Compatibility

**Decision:** Target NVDA 2025.1+ with lastTested 2025.3.2
**Date:** December 2025
**Status:** Updated

**Rationale:**
- Always target latest stable NVDA release
- User's system runs NVDA 2025.3.2
- Ensures access to newest APIs and improvements
- NVDA 2025.1 adds IUIAutomation6, improved speech, Remote Access

**Version Policy:** Update minimum and lastTested versions when new stable NVDA releases become available.

---

### 4. No Legacy Comment Support

**Decision:** Only support Modern Comments (PowerPoint 365)
**Date:** December 2025
**Status:** Final

**Rationale:**
- Legacy comments use different COM API
- Modern comments are the future
- Simplifies implementation significantly
- Target users are on 365

---

### 5. manifest.ini Quoting Format

**Decision:** Use specific quoting rules for manifest.ini
**Date:** December 2025
**Status:** Final - LEARNED THE HARD WAY

**The Rules:**
- No quotes for single words: `name = addonName`
- Double quotes for text with spaces: `summary = "My Addon"`
- Triple quotes for multi-line: `description = """Long text"""`
- No quotes for versions/URLs: `version = 0.1.0`

**Why This Matters:**
- Incorrect quoting causes NVDA to silently reject the addon
- Error messages are not helpful
- Took significant debugging time to discover

**Common Failures:**
- `summary = My Addon Name` → FAILS (needs quotes)
- `name = "addonName"` → May work but incorrect
- Using smart quotes (""") instead of straight quotes (""") → FAILS

---

### 6. Extend Built-in PowerPoint Support (Don't Replace)

**Decision:** Import built-in module and inherit from its AppModule class
**Date:** December 2025
**Status:** Updated - CRITICAL FIX

**Rationale:**
- NVDA has ~1500 lines of existing PowerPoint support
- Replacing it would break working features
- Extending allows adding comments without losing existing functionality

**WRONG Pattern (causes addon to not load):**
```python
# DON'T DO THIS - creates namespace collision
from nvdaBuiltin.appModules.powerpnt import *
class AppModule(appModuleHandler.AppModule):  # Wrong base class!
```

**CORRECT Pattern:**
```python
# appModules/powerpnt.py
from nvdaBuiltin.appModules import powerpnt as builtinPowerpnt

class AppModule(builtinPowerpnt.AppModule):  # Inherit from built-in
    # Add comment features on top
```

**Why This Matters:**
- The `import *` pattern imports the built-in `AppModule` into namespace
- Then defining our own `AppModule` with wrong base class causes confusion
- NVDA silently fails to load the addon - NO error in logs
- Fixed in v0.0.4 after testing showed addon wasn't initializing

---

### 7. Logging Strategy for Event Debugging

**Decision:** Use Python logging module extensively in Phase 1
**Date:** December 2025
**Status:** Final

**Rationale:**
- Events may not fire as expected
- Screen reader users cannot see console output
- NVDA log provides persistent debugging record
- Can verify behavior without visual feedback

**Implementation:**
```python
import logging
log = logging.getLogger(__name__)

log.debug("Event fired")
log.info("Important action")
log.error(f"Failed: {e}")
```

**View logs:** NVDA menu > Tools > View Log (NVDA+F1)

---

### 8. Testing Strategy - Manual First, Automated Later

**Decision:** Use manual NVDA testing for MVP, consider automation post-MVP
**Date:** December 2025
**Status:** Final

**Rationale:**
- Automated NVDA testing tools exist but are complex to set up
- Manual testing with scratchpad is fastest for iteration
- Real screen reader testing catches issues automation misses
- Automation useful for regression testing after MVP stable

**Manual Testing Workflow:**
1. Copy to scratchpad: `%APPDATA%\nvda\scratchpad\appModules\`
2. Enable scratchpad in NVDA settings
3. Reload plugins: NVDA+Ctrl+F3
4. Check NVDA log for errors

**Post-MVP Automation Options:**
- NVDA Testing Driver (C#)
- Guidepup (JavaScript)

---

## Backlogged Decisions

### Comment Resolution Status (Deferred)

**Issue:** Cannot reliably detect resolved vs unresolved comments
**Status:** Backlogged

**Why Deferred:**
- Resolved status not exposed in COM API
- Would require OOXML parsing (file is locked while open)
- Shadow copy approach is complex and fragile

**Research:** `research/PowerPoint-Comment-Resolution-LockedFile-Access-Research.md`

**Future Option:** Revisit if Microsoft exposes resolution status in COM API
