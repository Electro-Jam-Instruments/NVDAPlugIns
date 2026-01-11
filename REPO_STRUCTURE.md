# NVDA-Plugins Repository Structure

## Overview

Multi-plugin repository for NVDA accessibility addons developed by Electro Jam Instruments.

**GitHub Repo:** `Electro-Jam-Instruments/NVDAPlugIns`
**URL:** https://github.com/Electro-Jam-Instruments/NVDAPlugIns

## Directory Structure

```
NVDAPlugIns/
│
├── README.md                    # Repository index - lists all plugins
├── CHANGELOG.md                 # Repository-level changelog
├── CLAUDE.md                    # Repository-level dev instructions
├── REPO_STRUCTURE.md            # This file
├── LICENSE
│
├── .github/
│   └── workflows/
│       └── build-addon.yml      # Shared GitHub Actions workflow
│
├── .claude/
│   ├── commands/
│   │   └── deploy.md            # Deployment skill
│   └── skills/
│
├── docs/                        # Shared documentation & experts
│   └── experts/
│       └── expert-uia.md        # General UIA patterns (referenced by plugins)
│
├── powerpoint-comments/         # PowerPoint Comments Plugin
│   ├── README.md                # User-facing documentation
│   ├── CHANGELOG.md             # Plugin version history
│   ├── CLAUDE.md                # Plugin-specific dev instructions
│   ├── buildVars.py             # Plugin build configuration
│   ├── sconstruct               # Scons build script
│   ├── manifest.ini.tpl         # Addon manifest template
│   ├── manifest-translated.ini.tpl
│   ├── site_scons/              # Scons helpers
│   ├── addon/
│   │   ├── manifest.ini         # NVDA addon manifest
│   │   └── appModules/
│   │       └── powerpnt.py      # PowerPoint app module
│   ├── docs/                    # Plugin-specific docs
│   │   ├── experts/             # PPT-specific experts
│   │   ├── history/             # Development history
│   │   └── research/            # Research notes
│   └── tests/
│       └── resources/           # Test files (presentations, etc.)
│
├── windows-dictation-silence/   # Windows Voice Typing Silence Plugin
│   ├── README.md                # User-facing documentation
│   ├── CHANGELOG.md             # Plugin version history
│   ├── CLAUDE.md                # Plugin-specific dev instructions
│   ├── buildVars.py             # Plugin build configuration
│   ├── sconstruct               # Scons build script
│   ├── manifest.ini.tpl         # Addon manifest template
│   ├── manifest-translated.ini.tpl
│   ├── site_scons/              # Scons helpers
│   ├── addon/
│   │   ├── manifest.ini         # NVDA addon manifest
│   │   └── globalPlugins/
│   │       └── windowsDictationSilence.py
│   └── docs/                    # Plugin-specific docs
│
├── deletedocs/                  # Archived/deprecated documentation
└── localdocs/                   # Local-only documentation (not deployed)
```

## Plugin Types

| Plugin | Type | Scope |
|--------|------|-------|
| powerpoint-comments | AppModule | PowerPoint only |
| windows-dictation-silence | GlobalPlugin | System-wide |

## Build System

All plugins use the standard NVDA scons build system:

```bash
cd {plugin-folder}
scons
# Output: {plugin-name}.nvda-addon
```

Each plugin is self-contained with its own:
- `buildVars.py` - version and metadata
- `sconstruct` - build script
- `manifest.ini.tpl` - manifest template
- `site_scons/` - scons helpers

## Release Strategy

**Deployment:** Use `/deploy` command. See `.claude/commands/deploy.md`

### Tag Format (CRITICAL)

```
{plugin-name}-v{VERSION}[-beta]
```

Examples:
- `powerpoint-comments-v0.0.14-beta` - Beta release
- `powerpoint-comments-v1.0.0` - Stable release
- `windows-dictation-silence-v0.0.6-beta` - Beta release

**WARNING:** Tags without plugin prefix (e.g., `v0.0.1-beta`) will NOT trigger builds!

### Release Workflow

1. Update version in `{plugin}/buildVars.py`
2. Commit and push changes
3. Create and push tag: `git tag -a {plugin}-v{VERSION}-beta -m "Description"`
4. GitHub Actions builds and publishes automatically

### Download URLs

- Latest beta: `https://electro-jam-instruments.github.io/NVDAPlugIns/downloads/{plugin}-latest-beta.nvda-addon`
- Specific version: `https://electro-jam-instruments.github.io/NVDAPlugIns/downloads/{plugin}-{VERSION}.nvda-addon`

## Documentation Hierarchy

| Level | Location | Purpose |
|-------|----------|---------|
| Repository | `/CLAUDE.md` | Overview, links to plugins |
| Repository | `/docs/experts/` | Shared knowledge (UIA, etc.) |
| Plugin | `/{plugin}/CLAUDE.md` | Critical patterns for that plugin |
| Plugin | `/{plugin}/docs/` | Plugin-specific docs and experts |

## Plugin Independence

Each plugin is **self-contained**:
- Has its own README, CLAUDE.md, version
- Can be released independently
- Has its own documentation and tests
- No cross-plugin dependencies

Shared resources (`docs/experts/`) are optional conveniences.
