# Documentation Consolidation Plan

## Goals

1. **Consolidate knowledge** - Reduce large number of scattered files into directed docs for humans and agents
2. **Clear doc types** - Each directory should have clear purpose and document types
3. **Proper indexing** - READMEs at needed levels to support navigation
4. **Clean .agent/.claude** - Remove NVDA/PPT-specific content, leave correct startup references
5. **Expert agents** - Focused experts per domain that can be used as sub-agents

## Decisions Made

### File Naming Convention for Experts
- Pattern: `expert-{scope}-{domain}.md`
- Example: `expert-ppt-uia.md` - Expert file for PPT scope, UIA domain
- Rationale: Scope-first groups related files when sorted alphabetically
  - All PPT experts cluster together: `expert-ppt-*`
  - All NVDA experts cluster together: `expert-nvda-*`

### Source Folders Being Reviewed
- `.agent/experts/` - 6 expert folders with research
- `.claude/experts/` - 1 combined expert (nvda-ppt-developer.md)
- `.claude/knowledge/` - (if exists)

### Common Directory Structure (Applied at All Levels)

```
docs/
├── README.md                  # Index/navigation for this docs level
├── experts/                   # Expert agent definition files
│   └── expert-{scope}-{domain}.md
├── research/                  # Deep technical research (READ-ONLY reference)
│   └── {topic}-research.md
├── guides/                    # How-to guides (task-oriented)
│   └── {task}-guide.md
├── reference/                 # Reference docs (lookup-oriented)
│   └── {topic}-reference.md
├── history/                   # Historical docs (failed approaches, completed plans)
│   ├── pitfalls-to-avoid.md
│   └── failed-approaches.md
└── TODOs/                     # Active planning docs
    └── {feature}-plan.md
```

### Target Structure (Full Repository)

```
/                                   # Repository root
├── CLAUDE.md                       # Agent startup instructions (merged from .agent/agent.md)
├── docs/                           # Top-level NVDA core knowledge
│   ├── README.md                   # Index for general NVDA docs
│   ├── experts/
│   │   ├── expert-nvda-general.md  # General NVDA addon development expert
│   │   └── expert-nvda-testing.md  # NVDA testing patterns expert
│   ├── research/                   # General NVDA research (non-PPT specific)
│   ├── guides/
│   │   └── nvda-development-guide.md
│   └── reference/
│
├── powerpoint-comments/            # PPT plugin folder
│   ├── docs/                       # PPT-specific knowledge
│   │   ├── README.md               # Index for PPT docs
│   │   ├── experts/
│   │   │   ├── expert-ppt-uia.md   # UIA focus expert for PPT
│   │   │   └── expert-ppt-com.md   # COM automation expert for PPT
│   │   ├── research/               # PPT deep research
│   │   │   ├── PowerPoint-UIA-Research.md
│   │   │   ├── PowerPoint-COM-Research.md (consolidated)
│   │   │   └── ... (other PPT research)
│   │   ├── guides/                 # (existing)
│   │   ├── reference/              # (existing)
│   │   │   └── announcement-patterns.md (NEW - from bug fix notes)
│   │   ├── history/                # (existing)
│   │   │   ├── pitfalls-to-avoid.md (add 2 new pitfalls)
│   │   │   ├── failed-approaches.md
│   │   │   └── threading-refactor-plan.md (completed plan)
│   │   └── TODOs/
│   │       ├── feature-update-plan.md (existing)
│   │       └── docs-consolidation-plan.md (this file)
│   └── addon/                      # Plugin code
│
├── localdocs/                      # Local-only, NOT committed to git
│   └── vision/                     # AI vision research (future work)
│
├── deletedocs/                     # Staging for deletion review
│   └── approved/                   # Approved for deletion after final review
│
└── .claude/                        # Claude-specific config
    ├── experts/                    # (will be cleaned - PPT experts move to plugin docs)
    └── skills/                     # Skills remain here
```

## File-by-File Decisions

### windows-accessibility folder
| File | Decision | Target |
|------|----------|--------|
| windows-accessibility.md | MOVE + RENAME | `powerpoint-comments/docs/experts/expert-ppt-uia.md` |
| decisions.md | MERGE | Into `powerpoint-comments/docs/architecture-decisions.md` (verify relevance first) |
| research/PowerPoint-UIA-Research.md | MOVE | `powerpoint-comments/docs/research/` |

### nvda-plugins folder
| File | Decision | Target |
|------|----------|--------|
| nvda-plugins.md | SPLIT | General → `/docs/guides/nvda-development-guide.md`, PPT-specific → merge into PPT experts |
| decisions.md | MERGE | PPT decisions into PPT `architecture-decisions.md` |
| research/NVDA_PowerPoint_Native_Support_Analysis.md | MOVE | `powerpoint-comments/docs/research/` (PPT-specific) |
| research/NVDA-PowerPoint-Community-Addons-Research.md | MOVE | `powerpoint-comments/docs/research/` (PPT-specific) |
| research/NVDA_UIA_Deep_Research.md | MOVE | `docs/research/` (general NVDA) |
| research/PowerPoint-COM-Events-Research.md | MOVE | `powerpoint-comments/docs/research/` (PPT-specific) |
| slideshow-announcement-architecture.md | MOVE | `powerpoint-comments/docs/reference/` |

### accessibility-tester folder
| File | Decision | Target |
|------|----------|--------|
| accessibility-tester.md | SPLIT | General → `/docs/guides/nvda-testing-guide.md`, PPT tests → delete (outdated phase checklists) |

### nvda-addon-packager folder
| File | Decision | Target |
|------|----------|--------|
| nvda-addon-packager.md | MOVE | `deletedocs/approved/` (outdated - references old build system) |

### local-ai-vision folder
| File | Decision | Target |
|------|----------|--------|
| All files | MOVE | `localdocs/vision/` (keep local, future work) |

### powerpoint-automation folder
| File | Decision | Target |
|------|----------|--------|
| powerpoint-automation.md | MOVE + RENAME | `powerpoint-comments/docs/experts/expert-ppt-com.md` |
| decisions.md | MERGE | Into PPT `architecture-decisions.md` |
| research/*.md (6 files) | MOVE | `powerpoint-comments/docs/research/` (all PPT-specific) |

### Remaining .agent files
| File | Decision | Target |
|------|----------|--------|
| agent.md | MERGE | General info → root `CLAUDE.md`, remove PPT duplication |
| plans/threading-refactor-plan.md | MOVE | `powerpoint-comments/docs/history/` (completed plan) |
| context/project-history.md | SPLIT | AI vision parts → `localdocs/vision/`, rest delete |
| EXECUTION_PLAN_v0.0.1.md | MOVE | `deletedocs/approved/` (superseded) |
| normal-mode-announcement-order-fix.md | MOVE + RENAME | `powerpoint-comments/docs/reference/announcement-patterns.md` (active feature reference) |
| slideshow-notes-announcement-attempts.md | EXTRACT + DELETE | Add 2 pitfalls to pitfalls-to-avoid.md, then delete |

### .claude folder cleanup
| File | Decision | Target |
|------|----------|--------|
| experts/nvda-ppt-developer.md | DELETE | Replaced by focused experts in PPT docs |
| skills/readme-style.md | KEEP | Skills stay in .claude |

## Research Consolidation Notes

The research files need review and possible consolidation:
- Many overlap (e.g., multiple COM research files)
- Some may be outdated
- Target: Consolidate into fewer, comprehensive reference docs
- Naming: `{topic}-research.md` format

## Pitfalls to Add

From bug fix notes, add to pitfalls-to-avoid.md:

**Pitfall 13: self.currentSlide is NVDA Wrapper, Not COM Object**
- In overlay classes, `self.currentSlide` is an NVDA wrapper object
- Has NVDA properties (`appModule`, `TextInfo`) not COM properties (`NotesPage`, `SlideIndex`)
- Must access COM data via worker thread, not directly from overlay

**Pitfall 14: _get_name() Not Called in All Contexts**
- `_get_name()` works for normal mode slide focus
- In slideshow, NVDA uses `handleSlideChange()` → `reportFocus()` flow instead
- May need to override `reportFocus()` for slideshow customization

## Notes

### Scratchpad Testing
User noted they don't use scratchpad testing. Mark as optional advanced technique in testing guide.

### Verification Before Merge
Before merging decisions.md files:
1. Check current code implementation
2. Check existing docs in powerpoint-comments/docs/
3. Ensure not adding outdated information

## Status

- [x] Reviewed windows-accessibility folder
- [x] Reviewed nvda-plugins folder
- [x] Reviewed accessibility-tester folder
- [x] Reviewed nvda-addon-packager folder
- [x] Reviewed powerpoint-automation folder
- [x] Reviewed remaining .agent files
- [x] Reviewed bug fix notes for learnings
- [x] Defined expert agent naming convention (`expert-{scope}-{domain}.md`)
- [x] Defined common directory structure
- [x] Defined target repository structure
- [x] Execute consolidation - COMPLETED

## Execution Summary

**Completed Jan 2026:**
- Created `/docs/` folder structure with guides, research, experts, reference, history, TODOs
- Created `powerpoint-comments/docs/experts/` with expert-ppt-uia.md and expert-ppt-com.md
- Moved 10 PPT research files to `powerpoint-comments/docs/research/`
- Moved 1 general NVDA research file to `/docs/research/`
- Created nvda-development-guide.md and nvda-testing-guide.md from split files
- Merged 5 new decisions into architecture-decisions.md (now 16 total)
- Created announcement-patterns.md reference doc
- Added pitfalls 13-14 to pitfalls-to-avoid.md
- Moved local-ai-vision to localdocs/vision/
- Moved outdated files to deletedocs/approved/
- Updated CLAUDE.md with new docs structure
- Created README.md indexes at both docs levels
- Deleted old combined expert (.claude/experts/nvda-ppt-developer.md)

**Note:** Original .agent/ folder files are still in place - can be deleted after verifying everything works.
