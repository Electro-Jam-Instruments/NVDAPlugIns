# Spatial Typing Feedback - NVDA Addon

Moves the character echo out of the way, so typing stops colliding with what NVDA is reading.

## Status

**Status:** Active Development

## What you hear

| Where | What |
|-------|------|
| **Centre** | NVDA's voice, and each word as you complete it |
| **Right, quieter** | Each character as you type it, in a second voice |
| **Left, quieter** | The alert when you finish a misspelled word |

The centre is what you listen to. The two sides inform from the periphery without
competing for it.

### Example

**Typing `recieve ` in Word while NVDA reads the previous line:**

> Centre: the previous line continues uninterrupted
> Right, quieter: "r" "e" "c" "i" "e" "v" "e"
> Centre, on the space: "recieve"
> Left, quieter, on the space: the error alert

### The second voice

- Uses your installed Windows voices - the same ones NVDA offers
- Its own voice, speed, pitch and volume, independent of your main voice
- Rate boost available, so it can keep up with a fast main voice

## Tuning it

Everything is adjustable from NVDA's settings ring, so you hear each change as you make
it. `Ctrl+NVDA+Left/Right` to move between settings, `Ctrl+NVDA+Up/Down` to change them.

Pick a **stream** - main, characters or errors - then adjust that stream's voice, speed,
rate boost, pitch, volume, punctuation pauses and stereo position. Controls that do not
apply to a stream say so.

Changes to the character stream are spoken **by the character voice**, at its own
position and speed, so the announcement doubles as the sample.

Settings are saved when you leave the ring.

Also: `NVDA+shift+p` plays a test at each position, `NVDA+shift+s` toggles the add-on.

## Installation

1. Download the `.nvda-addon` file from [Releases](https://github.com/Electro-Jam-Instruments/NVDAPlugIns/releases)
2. Double-click to install
3. Restart NVDA when prompted

**Direct download:** [Latest Beta](https://electro-jam-instruments.github.io/NVDAPlugIns/downloads/spatial-typing-feedback-latest-beta.nvda-addon)

To hear the character stream you need **Speak typed characters** enabled in
NVDA → Preferences → Settings → Keyboard.

## Not compatible with Sound Split

NVDA's Sound Split changes channel volume for the whole NVDA process, above anything this
add-on does. If Sound Split is on, this add-on stays off and tells you so. Use one or the
other.

## Requirements

- NVDA 2025.1 or later
- Windows 11
- Windows voices installed
- A stereo output device - positioning has no effect on mono output

## Technical Details

- The character voice is a Windows `SpeechSynthesizer` the add-on activates itself, so it
  offers the same voices as NVDA without depending on NVDA's internals
- Positioning uses a constant-power pan law over per-channel volume
- Typed characters and completed words are separated at `speech.speakTypedCharacters`
- The error alert is redirected through the `nvwave.decide_playWaveFile` extension point

## Building

```bash
cd spatial-typing-feedback
scons
```

Output: `spatialTypingFeedback-X.X.X.nvda-addon`

## Documentation

See `docs/` - `architecture.md` for how it works and why, and
`research/nvda-and-winrt-reference.md` for the exact APIs it relies on.

## License

GNU General Public License v2.0 (GPL-2.0)

This addon is a derivative work of NVDA, which is licensed under GPL v2. See [LICENSE](../LICENSE) for full terms.

## Author

Electro Jam Instruments
