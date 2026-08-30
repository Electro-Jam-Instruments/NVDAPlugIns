# Spatial Typing Feedback - Next Steps

## Decisions (all settled)

- [x] **Plugin name.** `spatial-typing-feedback` / `spatialTypingFeedback` - confirmed.
- [x] **Scope of "typing error" (R2 / D5).** NVDA's typing-time spelling alert only (sound only - see D5's three-path table). **Caps-lock beep deferred** to a later version. Reading-time error reports (sound *and* the spoken "spelling error") deliberately stay with the main voice.
- [x] **Secondary voice engine (D3).** ~~SAPI 5~~ **OneCore**, via a second `ocSpeech` token, matching the main synth's voices and rate boost (R6). Decided - user is on OneCore for the rate boost. SAPI 5 is now only the fallback if S1 shows two tokens are not viable.
- [x] **Pan convention.** **-50 left / 0 centre / +50 right.** Main voice defaults to centre and is itself a pannable stream.
- [x] **Default secondary voice.** ~~Different voice~~ **Matches the main voice** - same OneCore voice, rate, rate boost and pitch. Separation is positional, not timbral (R6).
- [x] **Stream layout - three streams.** Main (main voice + completed words) 0 at 100% / Chars +35 at 50% / Errors -35 at 50%. Symmetric sides. Volume is per stream. Main and words share centre on purpose but stay separate audio paths, so they can be split later.
- [x] **When to build the settings layer.** Not until the values have been heard. Hardcode through S1/P1/P2, then freeze (D11).
- [x] **Name for the settings ring selector slot (R7 / D8).** **"Stream"** - confirmed.

## Listen first

- [ ] **Judge the values by ear.** `NVDA+shift+p` plays each stream. Type in Word with a deliberate misspelling. Is the character echo at +35 / 50% still intelligible under a boosted main voice? Does the error alert at -35 / 50% register? Edit `_streams.py` and `NVDA+control+F3` to iterate.
- [ ] **Then** freeze the values and build P3/P4 (D11).

## Spikes

- [ ] **S1 - BLOCKING, now testable in the real build: can we hold a second OneCore token?** Built and shipping; `NVDA+shift+p` proves it or the log says why. Call `ocSpeech_initialize` a second time while NVDA's own OneCore driver is live, speak through it, and confirm both voices work. The whole of D3 rests on this. If it fails, fall back to SAPI 5 and accept losing rate boost.
- [ ] **S1a - mono to stereo upmix.** OneCore returns mono; `setVolume(left=, right=)` needs a stereo player. Prove the upmix and confirm it is cheap enough at character-echo rates.
- [ ] **S1d - BLOCKING for R7: can the main voice be panned at all?** `getSynth()._player` is created with `channels=wav.getnchannels()` and OneCore is mono, so `setVolume(left=, right=)` may raise `E_INVALIDARG` on channel 1. Test whether WASAPI presents two channels anyway. If not, the main voice can only be centre and R7 degrades - the user must be told, not left with a control that does nothing.
- [ ] **S1c - remaining S1 questions.** Latency per character, WASAPI contention with the main synth's player, ducking interaction, mono output device degradation. See `../research/nvda-audio-and-typing-hooks.md` section 7.
- [ ] **S1b - overlapping utterances.** One space fires word echo and character echo together. Determine whether the secondary voice needs two `WavePlayer` instances (word centred, character right, playing at once) or whether queueing sounds acceptable.

## Phases

- [x] **P0 - Sound Split guard (D7).** If `soundSplitState` is not OFF, stay dormant and announce once. Mutually exclusive - no per-state logic. Must land with or before P1.
- [x] **P1 - panned error alert (R2).** Register on `nvwave.decide_playWaveFile`, match `textError.wav`, play panned left. Lowest risk, uses a supported extension point, useful on its own.
- [x] **P2 - panned typing echo (R1 and R3).** Wrap `speech.speakTypedCharacters` once, routing characters right and completed words centre. Must reproduce protected-typing, suppression-counter, terminal and `TypingEcho` mode behaviour exactly, and must fail safe.
- [ ] **P3 - settings panel (R5, R6, R8).** Only after the values stop moving (D11). Per-stream **pan and volume**. **Multi-stream master on/off**, OneCore voice selection from `ocSpeech_getVoices`, rate, **rate boost**, pitch, volume, per-channel pan and enable, plus the Sound Split notice from D7. Disable the rate boost control when `ocSpeech_supportsProsodyOptions()` is false.
- [ ] **P4 - settings ring slots (R7, D9).** Swap `globalVars.settingsRing` for a subclass that appends a **Stream** selector (Main / Chars / Errors) and a **Pan** slot (-50 to +50, reported as "left 50" / "centre"). Includes panning the main voice via `getSynth()._player`, which S1 must first prove is possible. No new gestures - NVDA's `Ctrl+NVDA+Arrow` keys reach it automatically.

## Future options (not scheduled)

- **3D HRTF positioning via OpenAL Soft.** Wanted, deliberately deferred. See D10 and `../research/spatial-audio-options.md`. Nothing in the current design blocks it.
- **Pan the caps-lock beep.** Deferred, and cheap when wanted: `tones.beep` already takes `left=` and `right=`, so it needs none of the wave-player machinery. Hook is `NVDAObject.event_typedCharacter`, gated on `config.conf["keyboard"]["beepForLowercaseWithCapslock"]`.
- **Speak the typing alert instead of thumping.** Say "error" in the secondary voice panned left rather than playing `textError.wav`. NVDA cannot do this today.

## Release plumbing

- [x] **Plugin card added to the downloads page.** `.github/workflows/release.yml` hardcodes the HTML for the GitHub Pages download page, including one `<div class="plugin-card">` per plugin and matching `loadChangelog` / `loadVersionInfo` calls in the inline script. A new plugin will not appear there until that HTML is edited. Do this as part of the first release, not before - the download buttons would point at files that do not exist.
- [x] **Added to the root README download table.**
