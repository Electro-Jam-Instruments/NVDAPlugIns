# Changelog

All notable changes to this repository will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## Plugin Changelogs

Each plugin maintains its own detailed changelog:

| Plugin | Version | Changelog |
|--------|---------|-----------|
| PowerPoint Comments | 0.1.0-beta | [powerpoint-comments/CHANGELOG.md](powerpoint-comments/CHANGELOG.md) |
| Windows Dictation Silence | 0.1.0-beta | [windows-dictation-silence/CHANGELOG.md](windows-dictation-silence/CHANGELOG.md) |

> **Note:** These are initial beta releases. Full stable releases coming soon.

## Repository Changes

Repository-level changes (build system, CI/CD, documentation structure) are tracked below.

### [1.0.0] - 2026-01-11

Initial public release of the NVDAPlugIns repository.

#### Added
- Multi-plugin repository structure
- GitHub Actions workflow for automated addon builds
- GitHub Pages hosting for downloads
- Shared documentation and expert knowledge base
- Standard NVDA scons build system for all plugins
- Contributing guidelines, security policy, code of conduct
- Issue and PR templates

#### Plugins Included
- **PowerPoint Comments** v0.1.0-beta - Comment navigation and slide notes for PowerPoint
- **Windows Dictation Silence** v0.1.0-beta - Auto-silence NVDA during Windows Voice Typing
