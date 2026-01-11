# README Skill for NVDA Addons

Use this skill when creating or updating README.md files for NVDA addons in this repository.

## Two Types of READMEs

### 1. User-Facing README (addon-folder/README.md)
- **Audience**: End users installing and using the addon
- **Tone**: Practical, task-oriented, shows what they'll hear
- **Focus**: Features, installation, keyboard shortcuts, examples

### 2. Developer Documentation README (addon-folder/docs/README.md)
- **Audience**: Developers maintaining or contributing to the addon
- **Tone**: Technical, reference-oriented, navigational
- **Focus**: Architecture, code structure, where to find things

---

## Tone and Language

### Critical: Avoid "Accessible"
- **DO NOT** use "Accessible" when describing features for screen reader users
- The features ARE accessible - that's the baseline expectation
- This addon UPDATES/ENHANCES the experience, it doesn't make it accessible
- Use instead: "Updated", "Enhanced", "Improved", "Cleaner"

### Action-Oriented Language
- Describe what the addon DOES, not what it IS
- Focus on user outcomes and workflows
- Be direct and practical

### Examples
- WRONG: "Accessible PowerPoint comments navigation"
- WRONG: "Makes PowerPoint comments accessible"
- RIGHT: "Updated PowerPoint comments and slide notes experience"
- RIGHT: "Cleaner comment reading with NVDA"

## README Structure

### 1. Title and Description
```markdown
# [App Name] [Feature] - NVDA Addon

[One sentence describing what the addon updates/improves] using NVDA screen reader.
```

### 2. Status Section
- Do NOT include version numbers (they go stale)
- Just show development status

```markdown
## Status

**Status:** Active Development
```

Valid statuses: "Active Development", "Beta", "Stable", "Maintenance"

### 3. Features Section
Organize by user task/workflow, NOT by technical implementation.

```markdown
## Features

### [User Task Category]
- **Feature name**: What it does

### [Another Task Category]

**Sub-workflow:**
- Step or shortcut
- Another step
```

Include "What You Hear" examples:
```markdown
### Example: What You Hear

**[Action taken]:**
> "Exact text NVDA speaks"
```

### 4. Installation Section
```markdown
## Installation

1. Download the `.nvda-addon` file from [Releases](https://github.com/Electro-Jam-Instruments/NVDAPlugIns/releases)
2. Double-click to install
3. Restart NVDA when prompted

**Direct download:** [Latest Beta](https://electro-jam-instruments.github.io/NVDAPlugIns/downloads/[addon-name]-latest-beta.nvda-addon)
```

### 5. Requirements Section
Only list actively supported platforms:
```markdown
## Requirements

- NVDA 2024.1 or later
- [Target Application and version]
- Windows 11
```

Note: Windows 10 is NOT supported (Microsoft ended support).

### 6. Technical Details (Optional)
Brief bullets about implementation approach - for developers:
```markdown
## Technical Details

- [Key technical approach]
- [Architecture note]
```

### 7. Building Section
```markdown
## Building

This addon uses the standard NVDA scons build system:

```bash
cd [addon-folder]
scons
```

Output: `[addonName]-X.X.X.nvda-addon`
```

### 8. Documentation Section (if docs/ folder exists)
```markdown
## Documentation

See the `docs/` folder for developer documentation:
- Architecture decisions
- Development guides
- Technical reference
```

### 9. License and Author
```markdown
## License

MIT License - See repository root for details.

## Author

Electro Jam Instruments
```

## Formatting Rules

### Keyboard Shortcuts
- Always use backticks: `F6`, `Ctrl+Alt+N`
- Sequential key presses use comma: `Alt, Z, C` (press Alt, then Z, then C)
- Simultaneous key presses use plus: `Ctrl+Alt+N` (hold all together)

### NVDA Output Examples
Use blockquotes for what user hears:
```markdown
> "Slide 3, has notes, has 2 comments, Project Timeline"
```

### Code/Text User Enters
Use code blocks:
```markdown
```
**** Your quick note here ****
```
```

## Template

```markdown
# [App] [Feature] - NVDA Addon

[Updated/Enhanced/Improved] [app] [feature] experience using NVDA screen reader.

## Status

**Status:** Active Development

## Features

### [Primary Feature Category]
- **Feature**: Description

### [Secondary Feature Category]

**[Sub-workflow]:**
- `Shortcut` - What it does
- Next step

### Example: What You Hear

**[Scenario]:**
> "What NVDA says"

## Installation

1. Download the `.nvda-addon` file from [Releases](https://github.com/Electro-Jam-Instruments/NVDAPlugIns/releases)
2. Double-click to install
3. Restart NVDA when prompted

**Direct download:** [Latest Beta](https://electro-jam-instruments.github.io/NVDAPlugIns/downloads/[name]-latest-beta.nvda-addon)

## Requirements

- NVDA 2024.1 or later
- [Application version]
- Windows 11

## Building

This addon uses the standard NVDA scons build system:

```bash
cd [folder]
scons
```

## License

MIT License - See repository root for details.

## Author

Electro Jam Instruments
```

---

## Developer Documentation README (docs/README.md)

This is for the `docs/` folder inside each addon. It helps developers navigate technical documentation.

### Structure

```markdown
# [Addon Name] - Developer Documentation

This folder contains technical documentation for developing and maintaining the [Addon Name] NVDA addon.

## Documentation Structure

### Top Level
- **architecture-decisions.md** - Key architectural decisions and patterns

### guides/
Development guides:
- **[guide-name].md** - Description

### reference/
Technical reference:
- **[reference-name].md** - Description

### history/
Lessons learned:
- **failed-approaches.md** - Approaches that didn't work and why
- **pitfalls-to-avoid.md** - Common mistakes to avoid

## When to Use Each Section

| If you need to... | Look at... |
|-------------------|------------|
| Understand why something was built a certain way | `architecture-decisions.md` |
| [Task] | `[file]` |
| Avoid repeating past mistakes | `history/` folder |
```

### Key Elements

1. **Clear folder structure** - List what's in each subfolder
2. **Navigation table** - "If you need to... look at..." format
3. **Brief descriptions** - One line per file explaining what it contains
4. **No user-facing content** - This is purely for developers
