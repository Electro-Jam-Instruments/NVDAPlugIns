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
