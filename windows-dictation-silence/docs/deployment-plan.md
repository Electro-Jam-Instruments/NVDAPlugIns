# Deployment Plan: Windows Dictation Silence Plugin

## Overview

Deploy the windows-dictation-silence NVDA plugin via GitHub Actions, mirroring the powerpoint-comments deployment pipeline.

## How PPT Plugin Deploys (Reference)

1. **Tag triggers workflow** - Push tag like `powerpoint-comments-v0.0.78-beta`
2. **Workflow validates** - Checks tag version matches manifest.ini
3. **Builds addon** - Runs `build-tools/build_addon.py`
4. **Creates GitHub Release** - Attaches `.nvda-addon` file
5. **Deploys to GitHub Pages** - Copies to downloads folder

The same workflow already supports multiple plugins - no workflow changes needed.

---

## Current State

| Item | Status |
|------|--------|
| Code version | v0.0.3 (keypress interception) |
| manifest.ini version | 0.0.1 (NEEDS UPDATE) |
| buildVars.py | Missing (optional) |
| GitHub Actions | Already configured |
| Build tools | Ready (`build-tools/build_addon.py`) |

---

## Pre-Deployment Steps

### Step 1: Update manifest.ini version

**File:** `windows-dictation-silence/addon/manifest.ini`

**Change:** `version = 0.0.1` → `version = 0.0.3`

### Step 2: Create buildVars.py (optional but recommended)

**File:** `windows-dictation-silence/buildVars.py`

```python
# -*- coding: UTF-8 -*-

addon_info = {
    "addon_name": "windowsDictationSilence",
    "addon_summary": "Auto-silence NVDA during Windows Voice Typing",
    "addon_description": """Automatically turns off NVDA speech when Windows Voice Typing (Win+H) is active.
Speech is restored when Voice Typing closes.""",
    "addon_version": "0.0.3",
    "addon_author": "Electro Jam Instruments",
    "addon_url": "https://github.com/Electro-Jam-Instruments/NVDAPlugIns",
    "addon_minimumNVDAVersion": "2024.1",
    "addon_lastTestedNVDAVersion": "2025.1",
}
```

---

## Deployment Steps

### Step 3: Commit changes

```bash
git add windows-dictation-silence/
git commit -m "windows-dictation-silence v0.0.3: Prepare for deployment"
```

### Step 4: Push to branch

```bash
git push origin PPTCommentReview
```

### Step 5: Create and push tag

```bash
git tag -a windows-dictation-silence-v0.0.3-beta -m "v0.0.3-beta: Auto-silence during Voice Typing"
git push origin windows-dictation-silence-v0.0.3-beta
```

### Step 6: Verify build

```bash
gh run list --workflow=build-addon.yml --limit 1
gh run watch
```

---

## CRITICAL: Tag Format

**CORRECT:** `windows-dictation-silence-v0.0.3-beta`

**WRONG:** `v0.0.3-beta` (missing plugin name - build won't trigger!)

The workflow trigger pattern is `*-v[0-9]+.[0-9]+.[0-9]+*` which REQUIRES the plugin name prefix.

---

## Expected Results

After successful deployment:

| Item | Value |
|------|-------|
| GitHub Release | `windows-dictation-silence-v0.0.3-beta` |
| Release Type | Pre-release (beta) |
| Addon File | `windows-dictation-silence-0.0.3.nvda-addon` |
| Download URL | `https://github.com/Electro-Jam-Instruments/NVDAPlugIns/releases/download/windows-dictation-silence-v0.0.3-beta/windows-dictation-silence-0.0.3.nvda-addon` |
| GitHub Pages | Updated with new plugin |

---

## Validation Checklist

Before tagging:
- [ ] manifest.ini version = 0.0.3
- [ ] buildVars.py version = 0.0.3 (if created)
- [ ] Code tested locally with NVDA scratchpad
- [ ] Changes committed and pushed

After tagging:
- [ ] Tag format correct: `windows-dictation-silence-v0.0.3-beta`
- [ ] Build workflow triggered (check GitHub Actions)
- [ ] Release created on GitHub
- [ ] Download link works
- [ ] Addon installs in NVDA

---

## Troubleshooting

### Build Didn't Run After Pushing Tag

**Cause:** Tag format is wrong (missing plugin name prefix)

**Fix:**
```bash
# Delete wrong tag
git tag -d v0.0.3-beta
git push origin :refs/tags/v0.0.3-beta

# Create correct tag
git tag -a windows-dictation-silence-v0.0.3-beta -m "v0.0.3-beta: Description"
git push origin windows-dictation-silence-v0.0.3-beta
```

### Version Mismatch Error

**Cause:** Tag version doesn't match manifest.ini version

**Fix:** Update manifest.ini version to match tag, commit, push, then create tag.

---

## Quick Deploy Command Sequence

```bash
# After code changes are ready:
VERSION="0.0.3"

# 1. Update manifest.ini version (manual or use bump script)
python build-tools/bump_version.py windows-dictation-silence $VERSION

# 2. Commit
git add windows-dictation-silence/
git commit -m "windows-dictation-silence v$VERSION: Prepare release"

# 3. Push
git push origin PPTCommentReview

# 4. Tag and push
git tag -a windows-dictation-silence-v${VERSION}-beta -m "v${VERSION}-beta: Auto-silence during Voice Typing"
git push origin windows-dictation-silence-v${VERSION}-beta

# 5. Watch build
gh run watch $(gh run list --workflow=build-addon.yml --limit 1 --json databaseId -q '.[0].databaseId')
```

---

## Future Releases

For subsequent releases, increment version and repeat:

```bash
VERSION="0.0.4"
python build-tools/bump_version.py windows-dictation-silence $VERSION
git add windows-dictation-silence/addon/manifest.ini
git commit -m "Bump windows-dictation-silence to v$VERSION"
git push origin PPTCommentReview
git tag -a windows-dictation-silence-v${VERSION}-beta -m "v${VERSION}-beta: Description"
git push origin windows-dictation-silence-v${VERSION}-beta
```
