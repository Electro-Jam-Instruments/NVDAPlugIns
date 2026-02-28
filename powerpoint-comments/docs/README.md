# PowerPoint Comments Addon - Developer Documentation

This folder contains technical documentation for the PowerPoint Comments NVDA addon.

## PowerPoint View Modes

The addon works in both PowerPoint view modes:

- **Normal view** - Where you edit slides (slide thumbnails on left, main slide in center, notes below)
- **Slideshow view** - Full-screen presentation mode (F5 to start)

Each mode has different NVDA object types and announcement patterns. See `reference/announcement-patterns.md` for Normal view and `reference/slideshow-override.md` for Slideshow view.

## Documentation Structure

| Folder | Purpose |
|--------|---------|
| `experts/` | Expert agent definition files for PPT-specific domains |
| `research/` | Deep technical research (PPT COM, UIA, etc.) |
| `guides/` | How-to guides for development tasks |
| `reference/` | API and pattern references |
| `history/` | Pitfalls, failed approaches, completed plans |
| `TODOs/` | Active planning documents |

## Key Documents

### Architecture
- **[architecture-decisions.md](architecture-decisions.md)** - All architectural decisions and patterns (16 decisions)

### Experts
- **[expert-ppt-uia.md](experts/expert-ppt-uia.md)** - UIA focus management expert (references `docs/experts/expert-uia.md` for general UIA)
- **[expert-ppt-com.md](experts/expert-ppt-com.md)** - COM automation expert

### Guides
- **complete-skeleton.md** - Full addon structure template
- **nvda-addon-development.md** - NVDA addon development patterns
- **powerpoint-com-api.md** - PowerPoint COM automation reference
- **testing-and-debugging.md** - How to test and debug

### Reference
- **[announcement-patterns.md](reference/announcement-patterns.md)** - Patterns for modifying NVDA announcements
- **comments-pane-reference.md** - PowerPoint comments pane structure
- **slideshow-override.md** - Slideshow mode content suppression
- **nvda-event-timing.md** - NVDA event timing and threading

### History
- **[pitfalls-to-avoid.md](history/pitfalls-to-avoid.md)** - Common mistakes (14 pitfalls documented)
- **failed-approaches.md** - Approaches that didn't work
- **threading-refactor-plan.md** - Completed threading architecture plan

### Planning
- **[MVP_IMPLEMENTATION_PLAN.md](MVP_IMPLEMENTATION_PLAN.md)** - Complete implementation plan for comment navigation

### Research
Deep dive research documents (9 files):
- PowerPoint-UIA-Research.md - UIA tree structure and patterns
- PowerPoint-COM-Automation-Research.md - Image extraction, tables, comments
- PowerPoint-COM-Events-Research.md - Event-driven slide detection
- NVDA_PowerPoint_Native_Support_Analysis.md - NVDA's built-in PowerPoint handling
- NVDA-Slideshow-Mode-Override-Research.md - TreeInterceptor override options
- Slideshow-Override-Plan-Review.md - Expert review of slideshow approach
- And more...

## When to Use Each Section

| If you need to... | Look at... |
|-------------------|------------|
| Understand why something was built a certain way | `architecture-decisions.md` |
| Work with UIA focus management | `experts/expert-ppt-uia.md` |
| Work with COM automation | `experts/expert-ppt-com.md` |
| Modify NVDA announcements | `reference/announcement-patterns.md` |
| Work on slideshow mode | `reference/slideshow-override.md` |
| Avoid repeating past mistakes | `history/pitfalls-to-avoid.md` |
| Deep dive on a topic | `research/` folder |
