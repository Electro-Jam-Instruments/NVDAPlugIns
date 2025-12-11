# Release Management

## Overview

This document defines the release process for NVDA plugins in this repository. All version updates are manual and controlled.

## Version Format

**Format:** `MAJOR.MINOR.PATCH` (e.g., `0.0.1`, `0.1.0`, `1.0.0`)

**Starting version:** `0.0.1`

| Component | When to Increment |
|-----------|-------------------|
| MAJOR | Breaking changes, major rewrites |
| MINOR | New features, significant enhancements |
| PATCH | Bug fixes, small improvements |

## Release Types

| Type | Tag Pattern | GitHub Release | Use For |
|------|-------------|----------------|---------|
| Beta | `pluginname-vX.X.X-beta` | Pre-release | Testing, validation |
| Release | `pluginname-vX.X.X` | Stable | Production-ready |

**Examples:**
- `powerpoint-comments-v0.0.1-beta` - First beta for testing
- `powerpoint-comments-v0.0.1` - First stable release
- `powerpoint-comments-v0.0.2-beta` - Next beta with fixes

## Current Versions

| Plugin | Current Version | Status |
|--------|-----------------|--------|
| powerpoint-comments | 0.0.13 | Beta - Phase 1 Complete |

## Version Update Process

**IMPORTANT: Version updates are manual and require explicit request.**

### Step 1: Update Version (Manual)

Run the bump script with the new version:

```bash
python build-tools/bump_version.py powerpoint-comments 0.0.2
```

This updates:
- `powerpoint-comments/addon/manifest.ini` (version field)

### Step 2: Commit Version Change

```bash
git add powerpoint-comments/addon/manifest.ini
git commit -m "Bump powerpoint-comments to v0.0.2"
```

### Step 3: Push Changes

```bash
git push origin main
```

### Step 4: Create and Push Tag

**For Beta:**
```bash
git tag powerpoint-comments-v0.0.2-beta
git push origin powerpoint-comments-v0.0.2-beta
```

**For Release:**
```bash
git tag powerpoint-comments-v0.0.2
git push origin powerpoint-comments-v0.0.2
```

### Step 5: GitHub Actions Auto-Build

When tag is pushed:
1. GitHub Actions triggers automatically
2. Builds .nvda-addon package
3. Creates GitHub release (pre-release for beta, stable for release)
4. Uploads addon file to release

## Download URLs

After release is created:

**Beta:**
```
https://github.com/Electro-Jam-Instruments/NVDAPlugIns/releases/download/powerpoint-comments-v0.0.1-beta/powerpoint-comments-0.0.1.nvda-addon
```

**Release:**
```
https://github.com/Electro-Jam-Instruments/NVDAPlugIns/releases/download/powerpoint-comments-v0.0.1/powerpoint-comments-0.0.1.nvda-addon
```

## GitHub Repository Setup (One-Time)

Before first release, configure GitHub:

1. **Enable GitHub Actions:**
   - Settings > Actions > General
   - Select "Allow all actions and reusable workflows"

2. **Set Workflow Permissions:**
   - Settings > Actions > General > Workflow permissions
   - Select "Read and write permissions"

3. **Verify .github/workflows/ exists:**
   - Must contain `build-addon.yml`

## Validation

The GitHub Action validates:
- Tag version matches manifest.ini version
- manifest.ini follows quoting rules
- Addon package builds successfully

**If validation fails:** Fix the issue, delete the tag, and re-tag after fixing.

## NVDA Version Requirements

NVDA will only load an addon if the version is different from what's installed. Always increment version before testing on a system that has the addon installed.

## Release Checklist

Before creating a release tag:

- [ ] All code changes committed
- [ ] Version updated in manifest.ini (via bump_version.py)
- [ ] Version commit pushed to main
- [ ] Tested locally with NVDA scratchpad
- [ ] Ready for beta or release tag

## Rollback Process

If a release has issues:

1. **Don't delete the release** (users may have downloaded it)
2. Fix the issue in code
3. Bump to next patch version
4. Create new release

## Troubleshooting

### Build Didn't Run After Pushing Tag

**Symptom:** You pushed a tag but no workflow appeared in Actions.

**Cause:** Tag format is wrong.

**Fix:**
1. Check your tag: `git tag -l | tail -5`
2. Tags MUST be: `pluginname-vX.X.X-beta` (e.g., `powerpoint-comments-v0.0.8-beta`)
3. NOT: `vX.X.X-beta` (missing plugin name prefix)

Delete wrong tag and recreate:
```bash
# Delete wrong tag
git tag -d v0.0.8-beta
git push origin :refs/tags/v0.0.8-beta

# Create correct tag
git tag -a powerpoint-comments-v0.0.8-beta -m "v0.0.8-beta: Description"
git push origin powerpoint-comments-v0.0.8-beta
```

### Installed Addon Version Doesn't Match

**Symptom:** NVDA shows old version even after "installing" new build.

**Possible Causes:**
1. Build never ran (check tag format above)
2. Downloaded cached/old file (clear browser cache)
3. NVDA needs restart after install

**Verify:**
```bash
# Check what version is on GitHub Pages
curl -sI "https://electro-jam-instruments.github.io/NVDAPlugIns/downloads/powerpoint-comments-latest-beta.nvda-addon" | grep Last-Modified

# Check installed addon version
type "%APPDATA%\nvda\addons\powerPointComments\manifest.ini" | findstr version
```

### Version Mismatch Between Files

Always keep these in sync:
- `powerpoint-comments/addon/manifest.ini` - version field
- `powerpoint-comments/buildVars.py` - addon_version field

The build validates tag version against manifest.ini only, but buildVars.py should match for consistency.

## Version History Log

| Date | Plugin | Version | Type | Notes |
|------|--------|---------|------|-------|
| 2025-12-10 | powerpoint-comments | 0.0.13 | beta | WORKING - Fixed COM access with comHelper (UIAccess privilege) |
| 2025-12-10 | powerpoint-comments | 0.0.12 | beta | Debug - Added INFO logging to track COM initialization |
| 2025-12-10 | powerpoint-comments | 0.0.11 | beta | WORKING - Deferred COM work with core.callLater (fixes speech) |
| 2025-12-10 | powerpoint-comments | 0.0.10 | beta | FAILED - super().event_appModule_gainFocus() crashed (method doesn't exist) |
| 2025-12-10 | powerpoint-comments | 0.0.9 | beta | PARTIAL - Module loads but speech blocked by COM work in event handler |
| 2025-12-10 | powerpoint-comments | 0.0.8 | beta | FAILED - Explicit alias import pattern did not work |
| 2025-12-10 | powerpoint-comments | 0.0.7 | beta | FAILED - Tag format wrong, build never ran |
| 2025-12-10 | powerpoint-comments | 0.0.6 | beta | Working build, wrong inheritance (base class) |
| 2025-12-10 | powerpoint-comments | 0.0.1-0.0.5 | beta | Initial development iterations |

### Key Learnings (v0.0.9-v0.0.13)

1. **AppModule Pattern:** Only `import *` then `class AppModule(AppModule):` works
2. **Event Handlers:** `event_appModule_gainFocus` is optional hook - NO super() call
3. **Blocking Events:** Heavy work in event handlers blocks NVDA speech - defer with `core.callLater()`
4. **COM Access:** Must use `comHelper.getActiveObject()` not direct `GetActiveObject()` due to UIAccess privileges

---

**Last Updated:** December 2025
