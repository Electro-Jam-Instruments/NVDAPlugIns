# NVDA Addon Packager - Sub-Agent Definition

## Purpose

Specialized agent for building and packaging NVDA addons (.nvda-addon files).

## When to Use

Use this sub-agent when:
- Creating a new .nvda-addon package
- Validating manifest.ini syntax
- Troubleshooting addon installation failures
- Managing version updates

## Reference Files

**Do not duplicate - reference these authoritative sources:**

| Topic | File | What It Contains |
|-------|------|------------------|
| Directory Structure | `REPO_STRUCTURE.md` | Plugin folder layout, manifest template |
| Build Commands | `build-tools/build_addon.py` | Actual build script |
| Version Updates | `build-tools/bump_version.py` | Version management script |
| Release Process | `RELEASE.md` | Full release workflow, tagging, GitHub Actions |
| GitHub Workflow | `.github/workflows/build-addon.yml` | Automated build on tag |

## Quick Reference

### Build Command
```bash
python build-tools/build_addon.py powerpoint-comments
```

### Version Update (Manual Only)
```bash
python build-tools/bump_version.py powerpoint-comments 0.0.2
```

### Why Manual Version Control?
- NVDA only loads addons when version changes
- Prevents accidental version bumps
- User controls when to increment

## manifest.ini Validation Rules

**This is the critical knowledge that causes most failures:**

| Field Type | Quote Style | Example |
|------------|-------------|---------|
| Single word (no spaces) | No quotes | `name = addonName` |
| Single line WITH spaces | `"double quotes"` | `summary = "My Addon"` |
| Multi-line text | `"""triple quotes"""` | `description = """Text"""` |
| Version/URL | No quotes | `version = 1.0.0` |

**Common Failures:**
- `summary = My Addon Name` → FAILS (needs quotes)
- `name = "addonName"` → May work but incorrect
- Using smart quotes ("") instead of straight quotes ("") → FAILS

## Required Manifest Fields

```ini
name = addonName
summary = "Description with spaces"
version = 0.0.1
author = "Name <email>"
minimumNVDAVersion = 2023.1
lastTestedNVDAVersion = 2024.4
```

For full template, see `REPO_STRUCTURE.md` > manifest.ini Template section.

## Troubleshooting

### Addon Won't Install
1. Check NVDA log for specific error (NVDA+F1)
2. Validate manifest.ini quoting (most common issue)
3. Ensure zip structure is correct (contents at root, no extra parent folder)

### "Invalid addon" Error
Almost always a manifest quoting issue. Check:
- Fields with spaces have double quotes
- Multi-line descriptions use triple quotes
- No smart quotes (copy/paste from Word can cause this)

### Addon Installs but Doesn't Work
1. Check NVDA log for import errors
2. Verify module file names match target app (e.g., `powerpnt.py` for PowerPoint)
3. Test with scratchpad first (see Accessibility Tester agent)

### Version Not Loading
NVDA caches addons by version. If version unchanged:
1. Run `bump_version.py` to increment
2. Rebuild and reinstall
3. Restart NVDA

## Validation Checklist

Before packaging:
- [ ] manifest.ini passes quoting validation
- [ ] All required manifest fields present
- [ ] No Python syntax errors in modules
- [ ] appModules named correctly for target app
- [ ] No __pycache__ directories (build script excludes these)
- [ ] No .pyc files (build script excludes these)

## Related Decisions

- **Decision 5:** manifest.ini quoting rules - see `nvda-plugins/decisions.md`
- **Decision 8:** Manual testing workflow - see `nvda-plugins/decisions.md`
