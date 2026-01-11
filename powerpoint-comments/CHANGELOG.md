# Changelog - PowerPoint Comments Plugin

All notable changes to the PowerPoint Comments plugin will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.0.79-beta] - 2025-01-11

Initial public beta release. Full stable release coming soon.

### Added
- **Comment count announcements** - Hear "has X comments" when changing slides
- **Slide notes detection** - Hear "has notes" for slides with meeting notes (marked with `****`)
- **Read notes shortcut** - Press Ctrl+Alt+N to hear slide notes
- **Comments pane navigation** - PageUp/PageDown to change slides while in Comments pane
- **Comment reformatting** - Cleaner "Author: comment" format instead of verbose default
- **Slideshow support** - Notes announced during presentations, full content reading suppressed

### Technical
- AppModule architecture extending NVDA's built-in PowerPoint support
- COM events (WindowSelectionChange, SlideShowNextSlide) for instant slide detection
- Background worker thread for non-blocking COM operations
- Custom overlay classes (CustomSlide, CustomSlideShowWindow) for name prefixing
- Custom TreeInterceptor to control slideshow announcements

---

## Version History (Development)

Pre-release development versions (internal testing):

| Version | Notes |
|---------|-------|
| 0.0.77-0.0.79 | Slideshow TreeInterceptor, first slide caching |
| 0.0.70-0.0.76 | Overlay class lazy _get_name() pattern |
| 0.0.56-0.0.69 | Slideshow mode support, notes detection |
| 0.0.44-0.0.55 | Comment reformatting, auto-tab to comments |
| 0.0.21-0.0.43 | COM events, multi-window support |
| 0.0.14-0.0.20 | Worker thread architecture |
| 0.0.9-0.0.13 | COM access fixes (comHelper) |
| 0.0.1-0.0.8 | AppModule inheritance pattern discovery |
