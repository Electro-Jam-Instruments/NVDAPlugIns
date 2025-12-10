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
│   └── common_buildVars.py      # Shared build configuration
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

### Tagging Convention

Each plugin uses its own tag prefix:

```
powerpoint-comments-v1.0.0      # Stable release
powerpoint-comments-v1.1.0-beta # Beta/pre-release
future-plugin-v2.0.0            # Different plugin
```

### Release Workflow

1. **Development** happens on `main` branch in plugin directory
2. **When ready to release:**
   - Update `CHANGELOG.md` in plugin directory
   - Update version in `buildVars.py` and `manifest.ini`
   - Create tag: `git tag powerpoint-comments-v1.0.0`
   - Push tag: `git push origin powerpoint-comments-v1.0.0`
3. **GitHub Actions** (optional) auto-builds and creates release
4. **Manual alternative:** Build locally, create GitHub release, upload asset

### Download URLs

Direct download links follow this pattern:

```
https://github.com/Electro-Jam-Instruments/NVDAPlugIns/releases/download/powerpoint-comments-v1.0.0/powerpoint-comments-1.0.0.nvda-addon
```

For latest stable (using GitHub API or redirect):
```
https://github.com/Electro-Jam-Instruments/NVDAPlugIns/releases/latest/download/powerpoint-comments.nvda-addon
```

Note: "Latest" only works if you want ONE plugin to be "latest" - for multi-plugin repos, use explicit version tags.

## Build Process

### Building a Single Plugin

```bash
cd powerpoint-comments
python ../build-tools/build_addon.py
# Output: powerpoint-comments-1.0.0.nvda-addon
```

### manifest.ini Template

```ini
name = powerpoint-comments
summary = Accessible PowerPoint Comment Navigation
description = Navigate and read PowerPoint comments with keyboard shortcuts and automatic announcements
author = Electro Jam Instruments
url = https://github.com/Electro-Jam-Instruments/NVDAPlugIns/tree/main/powerpoint-comments
version = 1.0.0
minimumNVDAVersion = 2023.1
lastTestedNVDAVersion = 2024.4
```

### buildVars.py Template

```python
addon_info = {
    "addon_name": "powerpoint-comments",
    "addon_summary": "Accessible PowerPoint Comment Navigation",
    "addon_description": "Navigate and read PowerPoint comments with keyboard shortcuts",
    "addon_version": "1.0.0",
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
