# Documentation Reorganization Status

## Current State (January 10, 2026)

### Completed
- [x] PPT plugin migrated to standard NVDA scons build system (v0.0.79)
- [x] PPT docs moved from `.agent/` and `.claude/knowledge/` to `powerpoint-comments/docs/`
- [x] Created staging folders for review before deletion

### PPT Plugin Doc Locations

**powerpoint-comments/docs/** (IN GIT - keeper docs):
- architecture-decisions.md
- guides/complete-skeleton.md
- guides/nvda-addon-development.md
- guides/powerpoint-com-api.md
- guides/testing-and-debugging.md
- reference/comments-pane-reference.md
- reference/mention-detection-research.md
- reference/nvda-event-timing.md
- reference/slideshow-override.md
- history/failed-approaches.md
- history/pitfalls-to-avoid.md

**powerpoint-comments/localdocs/** (LOCAL ONLY - gitignored):
- research/ (AI vision research - future feature)

**powerpoint-comments/deletedocs/** (STAGED FOR DELETION - need user review):
- agent/ (old .agent folder contents - 31 files)
- claude-knowledge/ (old .claude/knowledge contents)

### Root Level Cleanup Needed

**Empty folders to delete (after confirming empty):**
- `.claude/knowledge/` - contents moved to powerpoint-comments/docs/
- `.agent/` - contents moved to powerpoint-comments/deletedocs/
- Root `localdocs/` - if exists and empty

**Files to gitignore:**
- `**/localdocs/`
- `**/deletedocs/`

### Voice Plugin Build Migration
- Needs same scons setup as PPT plugin
- No docs to move (fresh plugin)

---

## Pending Actions

1. **Voice Plugin** - Apply scons build system
2. **Update .gitignore** - Add localdocs/ and deletedocs/ patterns
3. **Delete empty source folders** - After confirming contents moved
4. **User review of deletedocs/** - Before final deletion
5. **Update this doc** - Track what was in v0.1.0-release-cleanup.md that's still relevant

---

## Files Summary

| Location | Status | Action |
|----------|--------|--------|
| powerpoint-comments/docs/ | Committed | Keep in git |
| powerpoint-comments/localdocs/ | Untracked | Gitignore |
| powerpoint-comments/deletedocs/ | Untracked | User review, then delete |
| .agent/ | Deleted from git | Delete local folder |
| .claude/knowledge/ | Unknown | Check if empty, delete |
