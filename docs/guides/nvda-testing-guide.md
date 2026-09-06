# NVDA Addon Testing Guide

Patterns for testing NVDA addons and verifying accessibility workflows.

## Testing Strategy

**Manual First, Automation Later**

During MVP development, use manual NVDA testing. Consider automation only after MVP is stable.

## Manual Testing Methods

### Scratchpad Testing (Optional - Fastest Iteration)

1. Copy module to `%APPDATA%\nvda\scratchpad\appModules\`
2. Enable in NVDA: Settings > Advanced > Developer Scratchpad
3. Reload plugins: NVDA+Ctrl+F3
4. Check NVDA log for errors

### Full Addon Testing (Pre-Release)

1. Build .nvda-addon package (scons)
2. Install via double-click
3. Restart NVDA

### Remote Testing Workflow

For testing on separate systems via GitHub:

1. **Build & Release**
   - Push code to branch
   - Create tag: `git tag plugin-name-v0.0.1-beta`
   - Push tag: `git push origin plugin-name-v0.0.1-beta`
   - GitHub Actions builds and creates release

2. **On Test System**
   - Download .nvda-addon from GitHub Releases
   - Double-click to install
   - Restart NVDA
   - Test and document results

3. **Report Results**
   - Note NVDA version, app version, Windows version
   - Document what worked/failed
   - Check NVDA log for errors

## Log-Based Verification

Critical for debugging event firing:

```python
import logging
log = logging.getLogger(__name__)

# In your methods:
log.debug("Event fired: event_appModule_gainFocus")
log.info(f"View type detected: {view_type}")
log.error(f"Connection failed: {e}")
```

View logs: NVDA menu > Tools > View Log (or NVDA+F1)

## Test Checklist Template

```markdown
## Test Session: [Date]

### Environment
- NVDA Version:
- App Version:
- Windows Version:
- Addon Version:

### Tests Performed
- [ ] Addon installs without error
- [ ] App module loads (check log)
- [ ] Focus event fires (check log)
- [ ] [Add feature-specific tests]

### Issues Found
1.

### Notes

```

## Debugging Tips

### Event Not Firing?
1. Check NVDA log for errors
2. Verify class name in log output
3. Add log statements to track flow
4. Check if NVDA is overriding (see badUIAWindowClasses)

### COM Connection Fails?
1. Verify target app is running
2. Check for correct ProgID
3. Try manual connection in Python REPL first
4. Look for COM security issues

### Speech Not Happening?
1. Verify ui.message() is called (add log before)
2. Check speech priority
3. Ensure no exceptions before ui.message()
4. Test with simple "ui.message('test')" first

## What Runs Automatically (CI)

`.github/workflows/tests.yml` runs on every push and pull request. It needs no NVDA,
no Windows and no audio device, so it finishes in seconds:

| Job | What it catches |
|-----|-----------------|
| **static** | `tools/check_addons.py` - a module that no longer parses, a function deleted while its callers stayed behind, two interchangeable engines whose signatures drifted, a version the changelog never mentions |
| **unit** | The arithmetic: pan law, mono-to-stereo upmix, SSML escaping, rate mapping, config spec. Python 3.11 and 3.13 (3.13 is what NVDA 2026 ships) |
| **build** | Every add-on still packages, and nothing stale (`.pyc`, `__pycache__`) got swept into the archive |

Each check is there because the thing it checks for actually happened, and in every case
the first sign of it was the user's screen reader misbehaving.

**These tests are mutation-tested.** Break the pan law, drop the ampersand escape, raise
the prosody ceiling past the engine's saturation point or silence one channel of the
upmix, and the suite fails. That check is worth repeating when adding tests - the
saturation test originally compared a constant against itself and passed happily on
broken code.

```bash
python tools/check_addons.py
python -m pytest spatial-typing-feedback/tests -v
```

## The Limits of Headless Testing

CI can test logic and wiring. It cannot test this:

- **No audio device.** A GitHub-hosted runner has no sound hardware and the Windows audio
  service is not running. `nvwave.WavePlayer` cannot open a device, and per-channel
  `setVolume` - the entire mechanism behind stereo positioning - has nothing to act on.
  A virtual audio driver can be installed, but it needs signing and a reboot.
- **Voices are not guaranteed.** OneCore voice availability varies by Windows image and
  by locale.
- **Time-dependent faults do not appear.** The `0xc0000409` crash in Spatial Typing
  Feedback took hours of real typing to surface, because a use-after-free only bites once
  the freed block is reused. No CI run reproduces that.

So the split is: **CI proves the arithmetic and the wiring; a real machine proves the
sound.** Do not try to move the second half into CI - the cost is high and the result
still would not be evidence about what the add-on sounds like.

### Running NVDA headlessly, if you need to

NVDA's own system tests are the reference implementation: Robot Framework driving NVDA
launched against a scratch config profile, with a spy global plugin that captures speech
into a list instead of speaking it. The same shape works for an add-on:

1. Build the add-on and install it into a throwaway config directory
2. Launch NVDA with `--config-path` pointing at that directory
3. Capture what NVDA *decided* to say, rather than listening to it
4. Assert on the captured sequences, and on the NVDA log

The existing tools for this are listed below. Both drive a real NVDA on a real Windows
desktop - neither removes the need for a machine with a sound card.

### Testing speech without NVDA at all

For modules that only touch the speech engine, the cheapest harness is no screen reader
at all. `spatial-typing-feedback/tests/conftest.py` does this: it stubs `logHandler`,
`config` and `nvwave`, and exposes the add-on's source directory as a synthetic package
so the leaf modules import each other without executing the package `__init__.py` (which
would drag in most of NVDA). Copy that pattern for any add-on whose logic is separable
from its NVDA integration.

## Automated Testing (Post-MVP)

**Consider automation when:**
- MVP is stable and feature-complete
- Manual regression testing becomes time-consuming
- Need to test across multiple NVDA versions

### Options

#### NVDA Testing Driver
- GitHub: github.com/kastwey/nvda-testing-driver
- C# library for programmatic NVDA control
- Can verify speech output

#### Guidepup
- GitHub: github.com/guidepup/guidepup
- JavaScript library for screen reader automation
- Cross-platform (NVDA, JAWS, VoiceOver)

## Best Practices

1. **Test with actual screen reader** - keyboard-only testing misses speech issues
2. **Test all keyboard shortcuts** - verify announcements at each step
3. **Test error states** - what happens when app closes mid-use?
4. **Test with real content** - use actual documents
5. **Document expected behavior** - write down what NVDA should say
