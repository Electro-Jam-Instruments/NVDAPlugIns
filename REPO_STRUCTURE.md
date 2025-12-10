# NVDA-Plugins Repository Structure

## Overview

This document defines the multi-plugin repository structure for all NVDA accessibility plugins developed by Electro Jam Instruments.

## Repository Name

**GitHub Repo:** `Electro-Jam-Instruments/NVDAPlugIns`
**URL:** https://github.com/Electro-Jam-Instruments/NVDAPlugIns

## Directory Structure

```
NVDA-Plugins/
│
├── README.md                    # Repository index - lists all plugins
├── LICENSE                      # MIT or GPL (NVDA compatible)
├── CONTRIBUTING.md              # Contribution guidelines
├── .github/
│   └── workflows/
│       └── build-addon.yml      # Shared GitHub Actions workflow
│
├── build-tools/                 # Shared build utilities
│   ├── build_addon.py           # Script to package .nvda-addon files
│   └── bump_version.py          # Script to update version in manifest.ini
│
├── powerpoint-comments/         # PowerPoint Comments Plugin
│   ├── README.md                # Plugin-specific documentation
│   ├── CHANGELOG.md             # Version history
│   ├── buildVars.py             # Plugin build configuration
│   ├── addon/
│   │   ├── manifest.ini         # NVDA addon manifest
│   │   ├── appModules/
│   │   │   └── powerpnt.py      # PowerPoint app module
│   │   ├── globalPlugins/
│   │   │   └── pptCommentNav.py # Global plugin (if needed)
│   │   └── locale/              # Translations (future)
│   │       └── en/
│   │           └── LC_MESSAGES/
│   └── tests/                   # Plugin-specific tests
│       └── test_comment_detection.py
│
├── future-plugin-2/             # Template for next plugin
│   ├── README.md
│   ├── CHANGELOG.md
│   ├── buildVars.py
│   ├── addon/
│   │   ├── manifest.ini
│   │   ├── appModules/
│   │   └── globalPlugins/
│   └── tests/
│
└── test-resources/              # Shared test files
    ├── create_test_presentation.py
    └── Guide_Dogs_Test_Deck.pptx
```

## Release Strategy

**Full documentation: See [RELEASE.md](RELEASE.md)**

### Release Types

| Type | Tag Pattern | GitHub Release |
|------|-------------|----------------|
| Beta | `pluginname-vX.X.X-beta` | Pre-release |
| Release | `pluginname-vX.X.X` | Stable |

### Tagging Convention

```
powerpoint-comments-v0.0.1-beta  # Beta for testing
powerpoint-comments-v0.0.1       # Stable release
powerpoint-comments-v0.0.2-beta  # Next beta
```

### Release Workflow (Automated)

1. **Update version** (manual, only when requested):
   ```bash
   python build-tools/bump_version.py powerpoint-comments 0.0.1
   git add powerpoint-comments/addon/manifest.ini
   git commit -m "Bump powerpoint-comments to v0.0.1"
   git push origin main
   ```

2. **Create and push tag:**
   ```bash
   git tag powerpoint-comments-v0.0.1-beta  # or without -beta
   git push origin powerpoint-comments-v0.0.1-beta
   ```

3. **GitHub Actions automatically:**
   - Validates tag version matches manifest.ini
   - Builds .nvda-addon package
   - Creates GitHub release (pre-release for beta)
   - Uploads addon file

### Download URLs

After automated release:

```
https://github.com/Electro-Jam-Instruments/NVDAPlugIns/releases/download/powerpoint-comments-v0.0.1-beta/powerpoint-comments-0.0.1.nvda-addon
```

### Version Management

**IMPORTANT:** Version updates are manual and controlled.

- NVDA only loads addons when version changes
- Always bump version before testing on a system with addon installed
- Use `bump_version.py` script to update manifest.ini

## Build Process

### Building a Single Plugin

```bash
python build-tools/build_addon.py powerpoint-comments
# Output: powerpoint-comments/powerpoint-comments-0.0.1.nvda-addon
```

### manifest.ini Template

**CRITICAL: Quoting rules matter! See notes below.**

```ini
name = powerPointComments
summary = "Accessible PowerPoint Comment Navigation"
description = """Navigate and read PowerPoint comments with keyboard shortcuts and automatic announcements."""
author = "Electro Jam Instruments <contact@electrojam.com>"
url = https://github.com/Electro-Jam-Instruments/NVDAPlugIns/tree/main/powerpoint-comments
version = 0.0.1
minimumNVDAVersion = 2023.1
lastTestedNVDAVersion = 2024.4
```

**manifest.ini Quoting Rules:**
| Field Type | Quote Style | Example |
|------------|-------------|---------|
| Single word (no spaces) | No quotes | `name = addonName` |
| Single line WITH spaces | `"double quotes"` | `summary = "My Addon"` |
| Multi-line text | `"""triple quotes"""` | `description = """Text"""` |
| Version/URL | No quotes | `version = 1.0.0` |

**If NVDA rejects your addon, check quoting first!**

### buildVars.py Template

```python
addon_info = {
    "addon_name": "powerpoint-comments",
    "addon_summary": "Accessible PowerPoint Comment Navigation",
    "addon_description": "Navigate and read PowerPoint comments with keyboard shortcuts",
    "addon_version": "0.0.1",
    "addon_author": "Electro Jam Instruments",
    "addon_url": "https://github.com/Electro-Jam-Instruments/NVDAPlugIns",
    "addon_minimumNVDAVersion": "2023.1",
    "addon_lastTestedNVDAVersion": "2024.4",
}
```

## Plugin Independence

Each plugin is **self-contained**:
- Has its own README, CHANGELOG, version
- Can be released independently
- Has its own test suite
- No cross-plugin dependencies

Shared resources in `build-tools/` are for convenience only.

## GitHub Repository Settings

### Recommended Settings

1. **Branch protection** on `main`:
   - Require PR reviews for production releases
   - Allow direct pushes for development (optional)

2. **Release settings**:
   - Use tag-based releases
   - Mark beta versions as "pre-release"

3. **Topics/Tags** for discoverability:
   - `nvda`, `nvda-addon`, `accessibility`, `screen-reader`, `powerpoint`

## Migration from Current Structure

Current project location:
```
30 - A11Y PowerPoint NVDA Plug-In/
├── MVP_IMPLEMENTATION_PLAN.md
├── REPO_STRUCTURE.md (this file)
├── research/
└── test_resources/
```

When creating GitHub repo:
1. Create `NVDA-Plugins` repo on GitHub
2. Reorganize locally into `powerpoint-comments/` subdirectory
3. Move shared test resources to `test-resources/`
4. Push to GitHub

## Version History

| Date | Version | Notes |
|------|---------|-------|
| 2024-12-08 | 1.0 | Initial structure definition |
