# Execution Plan: v0.0.1-beta Release

**Created:** December 2025
**Branch:** PPTCommentReview (NOT main)
**Status:** In Progress

## Objective

Deploy first working v0.0.1-beta release of PowerPoint Comments NVDA addon from the PPTCommentReview branch.

## Branch Strategy

- All work on `PPTCommentReview` branch
- NO push to main
- Tag created from PPTCommentReview
- Merge to main: Later, when user is ready

## Tasks

### Part 1: Agent Reviews

| Agent | Status | Notes |
|-------|--------|-------|
| PowerPoint Automation | Pending | Add Reference Files table, cross-references |
| Windows Accessibility | Pending | Deep review + update |
| Local AI Vision | Pending | Mark as deferred, add references |

### Part 2: Create Phase 1 Plugin Files

```
powerpoint-comments/
├── addon/
│   ├── manifest.ini          # v0.0.1
│   └── appModules/
│       └── powerpnt.py       # Phase 1 code
├── buildVars.py
└── README.md
```

### Part 3: Build & Deploy

1. [ ] Run `python build-tools/build_addon.py powerpoint-comments`
2. [ ] Verify .nvda-addon file created
3. [ ] Commit all files to `PPTCommentReview` branch
4. [ ] Push `PPTCommentReview` to GitHub
5. [ ] Create tag `powerpoint-comments-v0.0.1-beta` from `PPTCommentReview`
6. [ ] Push tag to trigger GitHub Actions
7. [ ] Monitor Actions workflow
8. [ ] Verify release created

### Part 4: Final Deliverable

- Link to v0.0.1-beta release
- Download URL for .nvda-addon file
- Summary of what was done

## What NOT To Do

- Push anything to main branch
- Accessibility Tester agent work
- NVDA testing setup

## Expected Download URL

```
https://github.com/Electro-Jam-Instruments/NVDAPlugIns/releases/download/powerpoint-comments-v0.0.1-beta/powerpoint-comments-0.0.1.nvda-addon
```
