# Spatial Typing Feedback - NVDA Addon

Separates typing feedback from NVDA's main speech by stereo position and voice, so the two stop competing for your attention.

## Status

**Status:** Beta

## Features

### Three streams, three places

| Stream | What you hear |
|-------|---------------|
| **Main** (centre) | NVDA's main voice, and each word as you complete it |
| **Chars** (right, quieter) | Each character as you type it |
| **Errors** (left, quieter) | The alert when you finish a misspelled word |

Main is what you listen to. Chars and Errors sit the same distance out at the same
level, so they inform without competing.

- **Honours your settings**: uses NVDA's existing *Speak typed characters* and *Speak typed words* settings, including "edit controls only" and password field suppression

### Secondary voice
- **Same voice as your main one**: matching Windows OneCore voice, rate, pitch and rate boost - the streams are told apart by position, not by sounding different
- **Rate boost supported**: the echo runs at the same boosted rate as your main voice, so it never lags behind

### Example: What You Hear

**Typing `recieve ` in Word with both echoes on, while the main voice reads the previous line:**

> Centre: previous line continues uninterrupted
> Right, quieter: "r" "e" "c" "i" "e" "v" "e"
> Centre, on the space: "recieve"
> Left, quieter, on the space: error sound

### Adjusting on the fly
- **Uses NVDA's settings ring**: `Ctrl+NVDA+Left/Right` to reach the add-on's settings, `Ctrl+NVDA+Up/Down` to change them - the same keys you already use for rate and pitch
- **Two new slots**: one to pick which stream you're adjusting, one to set its stereo position on a familiar -50 / centre / +50 scale
- **Every stream moves**, including your main voice, which sits centre by default
- **One switch to turn it all off**: back to NVDA's normal behaviour instantly

### Adjusting on the fly

Stream positions and levels are fixed in this build - they become adjustable once the
values have settled. Two commands are available now:

- `NVDA+shift+p` - play a test phrase at each stream position
- `NVDA+shift+s` - toggle spatial typing feedback on and off

## Installation

1. Download the `.nvda-addon` file from [Releases](https://github.com/Electro-Jam-Instruments/NVDAPlugIns/releases)
2. Double-click to install
3. Restart NVDA when prompted

**Direct download:** [Latest Beta](https://electro-jam-instruments.github.io/NVDAPlugIns/downloads/spatial-typing-feedback-latest-beta.nvda-addon)

## Not compatible with Sound Split

NVDA's Sound Split changes channel volume for the whole NVDA process, above anything
this add-on does. If Sound Split is on, this add-on stays off and tells you so. Use one
or the other.

## Requirements

- NVDA 2024.1 or later
- Windows 11
- Windows OneCore voices (the secondary voice matches your main one)
- A stereo output device - positioning has no effect on mono output

## Technical Details

- GlobalPlugin architecture for system-wide operation
- Panning via `nvwave.WavePlayer.setVolume(left=, right=)`
- Secondary voice is a second independent Windows OneCore instance, so it offers the same voices and the same rate boost as NVDA's main synthesizer
- Audio synthesised to a buffer, upmixed to stereo and played through our own wave player, so NVDA's single-synthesizer model is left intact
- Character and word echo both intercepted at `speech.speakTypedCharacters`, where NVDA decides them
- Error sound redirected through the `nvwave.decide_playWaveFile` extension point

## Building

This addon uses the standard NVDA scons build system:

```bash
cd spatial-typing-feedback
scons
```

Output: `spatialTypingFeedback-X.X.X.nvda-addon`

## Documentation

See the `docs/` folder for developer documentation:
- User requirements
- Architecture decisions
- Research on the NVDA hooks this addon depends on

## License

GNU General Public License v2.0 (GPL-2.0)

This addon is a derivative work of NVDA, which is licensed under GPL v2. See [LICENSE](../LICENSE) for full terms.

## Author

Electro Jam Instruments
