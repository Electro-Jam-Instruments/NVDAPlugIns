# Spatial Typing Feedback - User Requirements

## Problem

NVDA speaks everything through one voice, in one place. Keyboard echo, typing-error
alerts, word-completion suggestions and normal document reading all arrive on the same
channel, in the same voice, competing for the same moment of attention. When you are
typing quickly the echo collides with whatever NVDA is reading, and there is no
pre-attentive cue telling you *which kind* of feedback you just heard.

Humans separate concurrent audio streams very well when the streams differ in **spatial
position** and **timbre**. This add-on gives each class of typing feedback its own
position in the stereo field and its own voice, so the streams stop competing.

## Requirements

### Stream layout - three streams: Main, Chars, Errors

| Stream | Contains | Pan | Volume |
|--------|----------|-----|--------|
| **Main** | NVDA's main voice **and** completed words | 0 | 100% |
| **Chars** | Typed characters | +35 | 50% |
| **Errors** | Typing errors | -35 | 50% |

Symmetric: the two side streams sit the same distance out and at the same level, and the
centre stream is what you actually listen to.

Main speech and word echo share the centre deliberately. They are different audio sources
under the hood - NVDA's own synth and our secondary voice - but they occupy one position,
so from the user's side there are three places, not four.

### R1 - Character echo on the right, in a secondary voice

**Given** NVDA's typed-character echo is turned on (*Speak typed characters*),
**when** the user types a character,
**then** that character is spoken by a *secondary* TTS voice, panned **+35 right at 50%
volume**, **and** NVDA's main voice does not also speak it.

Notes:
- Must honour NVDA's existing echo setting, not replace it. If character echo is off,
  nothing is redirected.
- Must honour "edit controls only" mode, the same way NVDA does.
- Must honour password/protected field suppression - protected fields never leak.

### R2 - Typing errors on the left

**Given** NVDA's typing-error alert is enabled,
**when** the user finishes a word that the application has flagged as a spelling error,
**then** the error alert is played **panned -35 left at 50% volume** - mirroring the
character echo on the other side.

Open question (see Decision D5): whether "error" here means only NVDA's existing
spelling-error alert, or should also cover other typing failures.

### R3 - Completed words in the centre, in the secondary voice

**Given** NVDA's typed-word echo is turned on (*Speak typed words*),
**when** the user completes a word - types a space or punctuation that ends it,
**then** the whole word is spoken by the *secondary* TTS voice, **centre at full
volume**, **and** NVDA's main voice does not also speak it.

Notes:
- This is NVDA's word echo, not application autocomplete. It fires on word completion:
  you type `h e l l o` then space, and hear "hello".
- Character echo (R1) and word echo (R3) are independent settings and can both be on.
  When both are on, pressing space produces the completed word centre at full volume
  *and* the space character right at half volume, at the same instant. That simultaneity
  is the point - the two are distinguishable by position *and* level.
- Same protected-field and "edit controls only" rules as R1.

### R4 - Nothing regresses when the feature is off

The add-on must be switchable off entirely, and when off NVDA behaves exactly as it
did before install. Uninstalling must leave no residue in NVDA's speech pipeline.

### R5 - Configurable

At minimum the user must be able to set:
- Secondary voice (voice + rate + rate boost + pitch + volume), independently of the
  main voice
- Pan position per stream, including the main voice, on a -50 to +50 scale
- **Volume per stream** - the character echo runs quieter than the rest by design
- Per-stream enable/disable
- A master multi-stream on/off (R8)

### R6 - The secondary voice matches the main voice

**Given** the user has chosen Windows OneCore as NVDA's synthesizer,
**then** the secondary voice uses the **same engine, the same voice and matching voice
settings** - including **rate boost** - as the main voice.

Separation between streams comes from **stereo position, not from timbre.** The two
voices sound the same; they are told apart by where they are.

Rationale:
- Rate boost is why the user is on OneCore. A secondary voice capped at normal OneCore
  speed would still be echoing while the main voice had moved on.
- A deliberately different voice was an earlier idea and is *not* what is wanted. One
  familiar voice in two places is easier to listen to than two voices competing.

Notes:
- Rate boost on OneCore requires `ocSpeech_supportsProsodyOptions()` (OneCore API > 5).
  Where unsupported, neither voice has it, and the control should be disabled rather
  than silently ignored.
- The user may still override the secondary voice independently (R5). Matching is the
  default, not a constraint.

### R7 - Adjustable from the synth settings ring

**Given** the user is at the keyboard with no dialog open,
**when** they use NVDA's normal settings ring gestures (`Ctrl+NVDA+Left/Right` to move
between settings, `Ctrl+NVDA+Up/Down` to change the value),
**then** they can reach and adjust every stream's position in the same ring as Rate,
Pitch and Volume - without opening a settings dialog.

Two new slots:

| Slot | Purpose | Values |
|------|---------|--------|
| **Stream** | Selects which stream the Pan slot acts on. Changing it produces no audio change by itself. | Main / Chars / Errors |
| **Pan** | Stereo position of the selected stream | **-50 (left) to +50 (right), 0 = centre** |

**Main is a stream too.** It is adjustable like any other, and its default position is
centre. Every stream must be settable from the ring, not just the secondary ones.

Pan scale is **-50 / 0 / +50** for left / centre / right - the familiar balance
convention - not -100 to +100.

Notes:
- The slot is called **Stream**. "Voice" was the first working name but collides with
  NVDA's existing Voice slot, which picks the synthesizer voice.
- The Stream slot scopes **only this add-on's slots**. NVDA's built-in Rate, Pitch,
  Volume and Voice slots continue to act on the main synthesizer regardless of what
  Stream is set to. This is a real trap and must be made obvious in how the slots are
  named and spoken.
- Scheduled after the core features work (phase P4).

### R8a - Never fail silently

**Given** any condition that prevents the stereo positions from working - Sound Split
being enabled, a mono output device, a failed hook,
**then** the add-on returns every stream to NVDA's normal behaviour and **says so once**,
naming the cause.

Rationale: the failure mode here is not cosmetic - a stream can end up inaudible rather
than merely mispositioned. A user who has lost keyboard echo and does not know why is
much worse off than one whose feature paused itself and said so. Degrading is not
acceptable; going dormant and explaining is.

### R8 - Multi-stream master on/off

**Given** the settings panel,
**then** there is a single **multi-stream on/off** control that enables or disables the
whole feature at once.

When off, every stream returns to NVDA's normal behaviour: the main voice unpanned,
echo through the main synth, error alert unpanned. This is the concrete implementation
of R4's "nothing regresses when the feature is off", and it is what the user reaches for
when the add-on is getting in the way rather than helping.

Notes:
- Belongs in the settings panel. Worth also exposing as a boolean slot in the settings
  ring (NVDA already supports boolean ring slots) so it can be toggled without opening a
  dialog - recommended, but the panel is the requirement.

### R9 - 3D positioning (future option, not scheduled)

Positioning today is **pan and volume only**. That is the shipping scope.

Kept on the books because it is wanted later: full 3D head-related transfer functions,
placing streams anywhere around the head rather than on the left-right line.

Nothing in the current design blocks it - we own the PCM buffer, so a future version can
transform the audio however it likes. Costs and honest limits are in
`docs/research/spatial-audio-options.md` and Decision D10.

## Non-goals (for now)

- Application autocomplete / suggestion lists (a different feature; NVDA reports those
  through ordinary object presentation, not through the typing echo path)
- Changing NVDA's main synthesizer
- Braille output changes
- Working alongside NVDA's own Sound Split - the two are mutually exclusive (D7)

## Success criteria

The user can type a sentence in Word with a deliberate misspelling and, without
pausing, hear: characters on the right in a distinguishable voice, each completed word
in the centre in that same voice, and an error thump on the left - all without the main
voice being interrupted or doubled up.
