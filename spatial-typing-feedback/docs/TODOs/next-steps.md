# What is left

Everything below is real remaining work. Anything settled has been removed rather than
ticked off - the design is described in `../architecture.md`.

## Correctness

- [x] **Respect NVDA's speech mode and sleep mode.** Done. Our audio bypasses speak(),
      so nothing stopped us automatically; setting NVDA to off or beeps left the
      character echo talking. Worse, it made a silenced main voice look like a broken
      add-on and cost an evening chasing the wrong cause - the speech mode check in
      NVDA's speak() sits AFTER filter_speechSequence, so our filter still sees every
      sequence while nothing is spoken, and nothing is logged.
- [x] **Do not free memory an async WinRT call is still using.** Done, and it was
      crashing NVDA. Three separate faults: the SSML string was freed the moment
      `SynthesizeSsmlToStreamAsync` returned; `_await` abandoned operations that outran
      its timeout, after which the caller released the buffer `ReadAsync` was writing
      into; and `_doSetVoice` leaked the voice collection on every successful change.
      Symptom was `nvda.exe` dying in `MSTTSEngine_OneCore.dll` with `0xc0000409` after a
      long session, with no clean shutdown and nothing in NVDA's log.
- [ ] **Confirm the crash is actually gone.** The fixes are sound, but the failure took
      hours to appear, so one quiet session does not prove anything. If it recurs, get the
      faulting module from Windows Event Viewer again - NVDA's own log will not show it.
- [ ] **Follow the output device.** `config.conf["audio"]["outputDevice"]` is read once
      when a player is created. Change NVDA's audio device and our streams keep playing to
      the old one.
- [ ] **Follow configuration profile switches.** NVDA changes voice and rate per
      application; our settings never re-apply.

## Making it feel like a proper add-on

- [ ] **Settings panel** under Preferences -> Settings. The ring is right for tuning by
      ear, but a panel is what makes it legible to someone who did not build it, and is
      the right home for per-stream enable, engine choice and the Sound Split notice.
- [ ] **Add-on help.** `addon_docFileName` is `None`, so "Add-on help" in the Add-ons
      Manager does nothing. Ship `doc/en/readme.html`.
- [ ] **Translations catalogue.** Strings are marked; there is no `locale/`.
- [ ] **First-run orientation.** It installs and silently changes how typing sounds.

## Before other people install it

- [ ] **Make replacing NVDA's ring slots optional.** Right for this user, invasive as a
      default - it removes NVDA's standard slots from everyone.
- [ ] **Tests for the arithmetic.** Pan law, mono-to-stereo upmix, SSML escaping and the
      rate mapping all run without NVDA and are exactly where a silent bug would hide.
- [ ] **Paths nobody has exercised**: a non-OneCore main synth, a mono output device, no
      Windows voices installed, 32-bit NVDA.

## Release

- [ ] **Confirm the minimum NVDA version.** Declared 2025.1 because
      `speech.extensions.filter_speechSequence` does not exist in 2024.1 - without it ring
      announcements would come through twice. Worth testing on 2025.1 rather than assuming.
- [ ] **Add a plugin card to the downloads page.** `.github/workflows/release.yml`
      hardcodes the GitHub Pages HTML; a new plugin will not appear until it is edited.
      Already done for this plugin - verify it renders on the first release.

## Wanted later

- **3D positioning.** OpenAL Soft with `ALC_SOFT_loopback`. Nothing in the current design
      blocks it; see the end of `../architecture.md` for the cost and the honest limits.
- **Pan the caps-lock beep.** `tones.beep` already takes `left` and `right`.
