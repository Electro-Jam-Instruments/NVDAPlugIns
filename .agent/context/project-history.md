# Project History

## Timeline

### Phase 1: Initial Exploration (Early December 2025)

**Goal:** AI-powered image descriptions for PowerPoint accessibility

**Work Done:**
- Researched Florence-2 and local vision models
- Analyzed Surface Laptop 4 hardware capabilities
- Created initial plugin skeleton with AI focus

**Outcome:** Realized comment navigation was higher priority need

---

### Phase 2: Pivot to Comment Navigation (December 2025)

**Key Decision:** Defer AI features, focus on comment accessibility

**Rationale:**
- Comments are inaccessible with NVDA (confirmed via research)
- Comment navigation provides immediate value
- AI can be added later as enhancement

**Work Done:**
- Deep research into NVDA PowerPoint support
- Analyzed COM vs UIA approaches
- Researched @mention detection
- Investigated comment resolution status (backlogged due to file locking)

---

### Phase 3: MVP Planning (December 2025)

**Deliverables:**
- `MVP_IMPLEMENTATION_PLAN.md` v3.0 - 6-phase implementation plan
- `REPO_STRUCTURE.md` - Multi-plugin repository structure
- Test presentation (Guide Dogs deck)
- `.agent/` knowledge base with expert files

**Technical Decisions:**
- Use comtypes (not pywin32)
- COM for data, UIA for focus
- Auto-switch to Normal view
- Automatic slide change detection

---

## Archive Contents

The `archive/` folder contains superseded content:

| Folder/File | What It Was | Why Archived |
|-------------|-------------|--------------|
| `NVDA PPT plug-in/` | Initial AI-focused plugin | Approach changed |
| `optimized vision model/` | AI model research | Deferred to post-MVP |
| `IMPLEMENTATION_PLAN.md` | Early plan | Replaced by MVP plan |
| `docs/` | Mixed research/docs | Research moved to `.agent/experts/` |
| `research/` | Original research folder | Consolidated into expert folders |

**Note:** Archive is gitignored - not pushed to repo.

---

## Key Pivots

### AI → Comments

**When:** Early December 2025
**Why:** User feedback - comments more urgent than image descriptions
**Impact:** Completely changed MVP scope

### Single Plugin → Multi-Plugin Repo

**When:** December 8, 2025
**Why:** Future-proofing for additional NVDA plugins
**Impact:** Changed repo structure, release workflow

---

## Lessons Learned

1. **Start with user's actual pain point** - Comments before AI
2. **Research before coding** - COM/UIA decision saved rework
3. **Document decisions** - Future sessions benefit from context
4. **Keep research** - Even backlogged features may return
