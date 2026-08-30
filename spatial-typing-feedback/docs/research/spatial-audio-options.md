# Research: going beyond left/right (R9)

Verified against `nvaccess/nvda` master, fetched 2026-08-30.

---

## 1. The enabling fact

**We own the PCM buffer, so nothing in NVDA limits how spatial we get.**

Decision D3 has us synthesising OneCore audio into a buffer we control, then feeding it
to our own `nvwave.WavePlayer`. NVDA only ever sees finished stereo samples. Whatever
transformation we apply between those two points is entirely ours - amplitude panning,
interaural delay, full head-related transfer function convolution.

`WavePlayer.setVolume(left=, right=)` is a convenience, not a ceiling. It is the *only*
positioning NVDA offers, but it is not the only positioning available to us.

---

## 2. What the platform gives us - and does not

### NVDA does not bundle numpy or scipy

`pyproject.toml` runtime dependencies in full:

```
bleak, comtypes, cryptography, pyserial, wxPython, configobj, requests,
url-normalize, schedule, fast-diff-match-patch, pycaw, rpyc, pywin32, nh3,
crowdin-api-client, fuzzysearch, markdown, lxml, mdx_truly_sane_lists,
markdown-link-attr-modifier, mdx-gh-links, l2m4m, pymdown-extensions,
pyphen, detect-secrets, regex
```

No numeric stack. Any array maths is pure Python unless we bundle something.

### `audioop` is gone

`.python-versions`: `cpython-3.13.13-windows-x86_64-none`, and
`requires-python = ">=3.13,<3.14"`.

Python 3.13 removed `audioop` (PEP 594). The obvious C-accelerated path -
`audioop.tostereo(fragment, width, lfactor, rfactor)` - is **not available** on current
NVDA. Do not design around it, even though it would be perfect.

### NVDA master is now 64-bit

`x86_64` in `.python-versions`. Earlier NVDA releases (2024.1, 2025.1 - our current
`minimumNVDAVersion` and `lastTestedNVDAVersion`) are 32-bit x86.

**Any bundled native DLL therefore needs both an x86 and an x64 build**, selected at
runtime, if we want to span those versions. This is a real packaging cost and it lands
squarely on the HRTF option below.

---

## 3. Tier 1 - interaural time and level difference (no dependencies)

The cheapest real improvement over amplitude panning. Two cues instead of one:

- **ILD** (level difference) - what `setVolume` already does.
- **ITD** (time difference) - the far ear hears the sound up to ~0.7 ms later. This is
  the dominant localisation cue below about 1.5 kHz, and speech lives there.

Adding ITD to ILD noticeably improves externalisation - the sound stops feeling like a
balance control and starts feeling like a direction.

### It costs essentially nothing in pure Python

Both operations avoid per-sample arithmetic:

```python
import array

def monoToStereoWithITD(mono: array.array, delaySamples: int, delayRight: bool):
    """Interleave mono into stereo, delaying one channel by whole samples."""
    n = len(mono)
    stereo = array.array("h", bytes(2 * 2 * (n + delaySamples)))
    if delayRight:
        stereo[0 : 2 * n : 2] = mono                      # left, no delay
        stereo[2 * delaySamples + 1 :: 2] = mono          # right, delayed
    else:
        stereo[2 * delaySamples :: 2] = mono
        stereo[1 : 2 * n : 2] = mono
    return stereo
```

Extended slice assignment on `array` runs in C. No Python-level loop, no multiply.
The level half is free - `WavePlayer.setVolume` does it in WASAPI.

At 22.05 kHz (OneCore's typical rate), 0.7 ms is ~15 samples, so whole-sample delay
resolution is adequate. Sub-sample interpolation is not worth the cost.

### Optional: head shadow

A one-pole lowpass on the far ear adds the third natural cue. This *does* need a
per-sample loop - roughly 4 400 iterations for a 200 ms character echo. Likely a few
milliseconds in CPython. Measure before committing; drop it if it costs more than it
buys.

### What Tier 1 delivers

Better externalisation and a more convincing left/right axis. **Still one axis.** No
front/back, no elevation, no distance.

---

## 4. Tier 2 - real HRTF via OpenAL Soft

For positions off the left-right line, we need head-related transfer functions:
per-direction impulse responses convolved with the mono signal.

Rolling our own is not viable - a 256-tap convolution over a 200 ms utterance is ~2.5
million multiply-accumulates, which is fine in C and far too slow in pure Python, and we
have no numpy.

### OpenAL Soft is the right library

- Single DLL, roughly 1 MB, LGPL (dynamic linking is fine)
- Ships with a built-in HRTF dataset - no data files to source or license
- 3D positioning is `alSource3f(source, AL_POSITION, x, y, z)`; the library handles the
  convolution
- Mature and widely deployed

### The loopback extension keeps us on NVDA's audio path

`ALC_SOFT_loopback` lets OpenAL render into a buffer we supply instead of opening its
own output device:

```
ocSpeech -> mono PCM
         -> alcLoopbackOpenDeviceSOFT + alcRenderSamplesSOFT (HRTF on)
         -> stereo PCM
         -> nvwave.WavePlayer.feed()
```

This matters. Without loopback, OpenAL would open its own device and we would lose
NVDA's output-device selection, ducking, sound volume and Sound Split behaviour. With
loopback we get HRTF processing *and* stay inside NVDA's audio path.

### Honest limits of generic HRTF

Non-individualised HRTFs are not equally good in all directions:

| Cue | Reliability |
|-----|-------------|
| Left/right azimuth | **Good** - the strongest effect |
| Front arc positions | **Reasonable** |
| Front vs behind | **Poor** - the cone of confusion; without head tracking, "behind" frequently collapses to "in front" |
| Elevation | **Weak** - depends heavily on individual ear shape |
| Distance | **Moderate** - level plus filtering plus early reflections reads as near/far |

So a **ring of azimuth positions plus a near/far axis** is realistic. Reliable
"above and behind you" is not, and should not be promised.

### Headphones are required

HRTF assumes each ear hears only its own channel. Over speakers, crosstalk destroys the
effect and the result sounds worse than plain panning.

**Many screen reader users wear a single earbud** to keep one ear on the room. That
breaks binaural rendering completely - one ear gets one channel of a two-channel illusion.

There must be a mode switch: full HRTF for headphone users, Tier 1 or plain panning
otherwise. This cannot be auto-detected reliably; it has to be a setting.

### Costs

- Bundling a native DLL, in **both x86 and x64** builds (section 2)
- LGPL compliance - dynamic linking, ship the licence, offer the source
- Add-on size goes from kilobytes to megabytes
- One more processing stage per utterance; latency must be measured for character echo
- OpenAL context and source lifetime management on top of what we already own

---

## 5. Recommendation

**Tier 1 first, Tier 2 as a separate, later decision.**

Tier 1 costs nothing, has no dependencies, adds no packaging burden, and delivers the
single most valuable improvement - the streams stop sounding like a balance knob. It can
land inside P2 rather than as its own phase.

Tier 2 doubles the add-on's packaging complexity and pulls in a native dependency across
two architectures. That is worth paying **if** the answer to "what does better mean" is
front/back or elevation or a full circle. It is not worth paying if the answer is
"it should sound outside my head" - Tier 1 covers that.

This is why R9 asks the question before the spike is scoped.

---

## 6. Open questions

- Which of the four meanings in R9 matters most: more azimuth positions, front/back,
  distance, or externalisation?
- Headphones, speakers, or a single earbud? A single earbud rules out Tier 2 entirely
  and makes Tier 1 pointless too - that case needs plain level separation.
- How many streams need distinct positions at once? Three (main, echo, errors) fit on
  one axis. Five or six do not.
- Is a *moving* position ever useful, or are fixed positions per stream enough?

## Sources

- [nvaccess/nvda source (master)](https://github.com/nvaccess/nvda)
- [OpenAL Soft](https://openal-soft.org/)
- [PEP 594 - Removing dead batteries from the standard library](https://peps.python.org/pep-0594/)
