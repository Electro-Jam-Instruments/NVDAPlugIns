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
| powerpoint-comments | 0.0.1 | Development |

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

## Version History Log

| Date | Plugin | Version | Type | Notes |
|------|--------|---------|------|-------|
| (pending) | powerpoint-comments | 0.0.1 | beta | Phase 1 - Foundation |

---

**Last Updated:** December 2025
