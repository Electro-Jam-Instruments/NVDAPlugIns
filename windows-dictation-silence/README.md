# Windows Dictation Silence - NVDA Add-on

Use voice typing without hearing everything echoed back.

## What It Does

When you use Windows Voice Typing (Win+H), NVDA normally reads back every word as it appears. This creates an annoying echo where you hear yourself repeated constantly.

This add-on fixes that:

- **Press Win+H** - NVDA goes silent automatically
- **Start dictating** - No more echo, just your voice
- **Press any key** - NVDA speech comes back instantly

Your previous speech settings are preserved. If you had NVDA in "beeps" mode, it returns to beeps. If you had it in normal speech mode, it returns to that.

## How to Use

1. Press `Win+H` to open Windows Voice Typing
2. Dictate your text - NVDA stays quiet
3. Press any key (like Escape, or just start typing) to close Voice Typing
4. NVDA speech is back

That's it. No settings to configure, no extra shortcuts to remember.

## Installation

**Download:** [Latest Version](https://electro-jam-instruments.github.io/NVDAPlugIns/downloads/windows-dictation-silence-latest-beta.nvda-addon)

1. Click the download link above
2. Open the downloaded file
3. Restart NVDA when prompted

## Requirements

- NVDA 2024.1 or later
- Windows 11
- Windows Voice Typing enabled (built into Windows)

## Why This Exists

Windows 11 includes Voice Typing, activated with Win+H. It's a handy way to dictate text in any application. However, NVDA announces each word as it appears in the text field, creating a distracting echo effect.

NVDA has a built-in setting that's supposed to help ("selective UIA event registration"), but it doesn't fully solve the problem. This add-on provides a simple, reliable fix.

## Version History

See [CHANGELOG.md](CHANGELOG.md) for release notes.

## Get Help

Visit [community.electro-jam.com](https://community.electro-jam.com) for support and discussion.

## License

MIT License
