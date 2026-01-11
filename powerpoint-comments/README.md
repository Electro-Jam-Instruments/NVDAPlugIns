# PowerPoint Comments - NVDA Add-on

Never miss a comment or forget your presenter notes again.

## What It Does

This add-on makes working with PowerPoint comments and notes much easier:

- **Hear comment counts** - When you move to a new slide, you'll hear "has X comments" so you know what needs attention
- **Navigate comments easily** - Move through comments with arrow keys, hear who wrote each one
- **Presenter note reminders** - Mark important notes and hear "has notes" during your slideshow so you don't forget key points

## How to Use

### Slide Change Announcements

Just navigate slides normally. When you land on a slide, you'll hear:
> "Slide 3, has notes, has 2 comments, Project Timeline"

This tells you the slide number, whether it has notes you marked, how many comments, and the slide title.

### Working with Comments

**Open the Comments pane:**
- Press `Alt`, then `Z`, then `C`

**Move focus to Comments:**
- Press `F6` to jump to the Comments pane

**Navigate comments:**
- Arrow Up/Down to move between comments
- You'll hear "Author name: comment text" for each one
- Page Up/Down to move between slides while staying in the Comments pane

### Quick Notes for Presenters

Add personal reminders to your slides that get announced during presentations. Great for cues like "pause here" or "ask for questions".

**Setting up a quick note:**
1. Open the Notes pane for any slide
2. Wrap your reminder with `****` markers

**Example:**
```
**** Remember to demo the new feature ****

Regular presenter notes go here...
```

**During your slideshow:**
- You'll hear "has notes" after the slide title
- Press `Ctrl+Alt+N` anytime to hear your quick note read aloud

## Installation

**Download:** [Latest Version](https://electro-jam-instruments.github.io/NVDAPlugIns/downloads/powerpoint-comments-latest-beta.nvda-addon)

1. Click the download link above
2. Open the downloaded file
3. Restart NVDA when prompted

## Requirements

- NVDA 2024.1 or later
- Microsoft PowerPoint 365
- Windows 11

## Known Issues

- Comments pane must be open to get accurate comment counts
- The add-on will automatically open the Comments pane when you switch to a slide with comments

## Version History

See [CHANGELOG.md](CHANGELOG.md) for release notes.

## Get Help

Visit [community.electro-jam.com](https://community.electro-jam.com) for support and discussion.

## License

MIT License
