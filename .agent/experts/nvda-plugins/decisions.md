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

**Decision:** Target NVDA 2023.1+ with lastTested 2024.4
**Date:** December 2025
**Status:** Final

**Rationale:**
- 2023.1 is widely deployed
- Avoids deprecated API concerns
- Matches current NVDA addon ecosystem norms

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
