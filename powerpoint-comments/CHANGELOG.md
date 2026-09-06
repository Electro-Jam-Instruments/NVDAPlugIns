# Changelog - PowerPoint Comments Plugin

All notable changes to the PowerPoint Comments plugin will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Changed - 0.1.1
- **Loads on NVDA 2026.x again.** The add-on declared `lastTestedNVDAVersion` 2025.1,
  so NVDA 2026 flagged it incompatible and refused to run it. Every NVDA symbol the
  add-on uses was verified against 2026.2 before raising the declaration.
  This one leans on `from nvdaBuiltin.appModules.powerpnt import *` and on
  `comtypes.client._events._AdviseConnection`, a private comtypes internal; both were
  checked explicitly. Verified statically - the symbols exist, the behaviour has not
  been exercised on 2026.2 yet.


## [0.1.0-beta] - 2026-01-11

First public beta release with complete feature set.

### Added
- **Comment count announcements** - Hear "has X comments" when changing slides
- **Slide notes detection** - Hear "has notes" for slides with meeting notes (marked with `****`)
- **Read notes shortcut** - Press Ctrl+Alt+N to hear slide notes
- **Comments pane navigation** - PageUp/PageDown to change slides while in Comments pane
- **Comment reformatting** - Cleaner "Author: comment" format instead of verbose default
- **Slideshow support** - Notes announced during presentations, full content reading suppressed

### Fixed
- Race condition where slide data was one slide behind (v0.0.80 fix)
- Direct COM queries from overlay class for accurate real-time data

### Technical
- AppModule architecture extending NVDA's built-in PowerPoint support
- COM events (WindowSelectionChange, SlideShowNextSlide) for instant slide detection
- Background worker thread for non-blocking COM operations
- Custom overlay classes (CustomSlide, CustomSlideShowWindow) for name prefixing
- Custom TreeInterceptor to control slideshow announcements

---

## Development History

Pre-release versions (internal testing):

| Version | Notes |
|---------|-------|
| 0.0.80 | Fixed race condition with direct COM queries |
| 0.0.77-0.0.79 | Slideshow TreeInterceptor, first slide caching |
| 0.0.70-0.0.76 | Overlay class lazy _get_name() pattern |
| 0.0.56-0.0.69 | Slideshow mode support, notes detection |
| 0.0.44-0.0.55 | Comment reformatting, auto-tab to comments |
| 0.0.21-0.0.43 | COM events, multi-window support |
| 0.0.14-0.0.20 | Worker thread architecture |
| 0.0.9-0.0.13 | COM access fixes (comHelper) |
| 0.0.1-0.0.8 | AppModule inheritance pattern discovery |
