# Accessibility Tester - Sub-Agent Definition

## Purpose

Specialized agent for testing NVDA addons and verifying accessibility workflows.

## When to Use

Use this sub-agent when:
- Testing addon functionality with NVDA
- Verifying screen reader announcements
- Debugging event handling
- Planning test strategies

## Reference Files

| Topic | File | What It Contains |
|-------|------|------------------|
| Test Resources | `test-resources/` | Test presentations, scripts |
| Test Presentation | `test-resources/Guide_Dogs_Test_Deck.pptx` | PowerPoint with comments for testing |
| Create Test Data | `test-resources/create_test_presentation.py` | Script to generate test presentations |
| Testing Decision | `.agent/experts/nvda-plugins/decisions.md` | Decision 8: Manual First strategy |
| Phase Checklists | `MVP_IMPLEMENTATION_PLAN.md` | Test checklists per phase |

## Testing Strategy

**Decision 8: Manual First, Automated Later**

During MVP development, use manual NVDA testing. Consider automation only after MVP is stable and for regression testing.

### Manual Testing (Primary for MVP)

#### Scratchpad Testing (Fastest Iteration)
1. Copy module to `%APPDATA%\nvda\scratchpad\appModules\`
2. Enable in NVDA: Settings > Advanced > Developer Scratchpad
3. Reload plugins: NVDA+Ctrl+F3
4. Check NVDA log for errors

#### Full Addon Testing (Pre-Release)
1. Build .nvda-addon package
2. Install via double-click
3. Restart NVDA

### Remote Testing Workflow (Phase 1.1+)

For testing on separate systems via GitHub:

1. **Build & Release**
   - Push code to main branch
   - Create tag: `git tag powerpoint-comments-v0.0.1-beta`
   - Push tag: `git push origin powerpoint-comments-v0.0.1-beta`
   - GitHub Actions builds and creates release

2. **On Test System**
   - Download .nvda-addon from GitHub Releases
   - Double-click to install
   - Restart NVDA
   - Test and document results

3. **Report Results**
   - Note NVDA version, PowerPoint version, Windows version
   - Document what worked/failed
   - Check NVDA log for errors

### Log-Based Verification

Critical for debugging event firing:

```python
import logging
log = logging.getLogger(__name__)

# In your methods:
log.debug("Event fired: event_appModule_gainFocus")
log.info(f"View type detected: {view_type}")
log.error(f"COM connection failed: {e}")
```

View logs: NVDA menu > Tools > View Log (or NVDA+F1)

## Test Strategy by Phase

### Phase 1: Foundation + View Management

| Test Case | Method | Expected Result |
|-----------|--------|-----------------|
| App module loads | Check log | "PowerPoint Comments addon initialized" |
| Focus event fires | Check log | "event_appModule_gainFocus fired" |
| COM connection | Check log | "Connected to PowerPoint COM" |
| View detection | Check log | "View type: 9" (Normal) |
| View auto-switch | Manual | NVDA announces "Switched to Normal view" |

### Phase 1.1: Package + Deploy Pipeline

| Test Case | Method | Expected Result |
|-----------|--------|-----------------|
| Build script runs | Run build_addon.py | .nvda-addon file created |
| GitHub release created | Push tag | Release appears with download |
| Remote install works | Install on test system | Addon loads, log shows init message |

### Phase 2: Slide Change Detection

| Test Case | Method | Expected Result |
|-----------|--------|-----------------|
| Slide change detected | Change slides, check log | "Slide changed: 1 -> 2" |
| Comment count announced | Manual | NVDA speaks "3 comments on this slide" |
| No comments announced | Manual | NVDA speaks "No comments" |

### Phase 3: Focus First Comment

| Test Case | Method | Expected Result |
|-----------|--------|-----------------|
| Comments pane opens | Manual | Comments pane visible |
| Focus moves to comment | Manual | NVDA reads first comment |
| No comments handling | Manual | "No comments on this slide" |

### Phase 3.1: Slide Navigation from Comments

| Test Case | Method | Expected Result |
|-----------|--------|-----------------|
| Navigate to next slide | Keyboard shortcut | Moves to next slide, announces comments |
| Navigate to previous slide | Keyboard shortcut | Moves to previous slide, announces comments |
| Focus returns to comments | After navigation | Comments pane still focused |

### Phase 4: @Mention Detection

| Test Case | Method | Expected Result |
|-----------|--------|-----------------|
| @mention found | Open file with mentions | NVDA announces mention count |
| Current user detected | Check for own username | Correctly identifies user's mentions |
| No mentions | File without @mentions | No false positives |

### Phase 5: Polish + Packaging

| Test Case | Method | Expected Result |
|-----------|--------|-----------------|
| Error handling | Force errors | Graceful failure, user-friendly messages |
| Settings work | Change settings | Settings persist |
| Final build clean | Fresh install | No debug messages in release |

## Test Checklist Template

```markdown
## Test Session: [Date]

### Environment
- NVDA Version:
- PowerPoint Version:
- Windows Version:
- Addon Version:
- Test Presentation: Guide_Dogs_Test_Deck.pptx

### Tests Performed
- [ ] Addon installs without error
- [ ] App module loads (check log)
- [ ] Focus event fires (check log)
- [ ] COM connects (check log)
- [ ] [Add phase-specific tests]

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
1. Verify PowerPoint is running
2. Check for correct ProgID ("PowerPoint.Application")
3. Try manual connection in Python REPL first
4. Look for COM security issues

### Speech Not Happening?
1. Verify ui.message() is called (add log before)
2. Check speech priority
3. Ensure no exceptions before ui.message()
4. Test with simple "ui.message('test')" first

## Automated Testing (Post-MVP)

**Trigger Criteria:** Consider automation when:
- MVP is stable and feature-complete
- Manual regression testing becomes time-consuming
- Need to test across multiple NVDA/PowerPoint versions

### Options

#### NVDA Testing Driver
- GitHub: github.com/kastwey/nvda-testing-driver
- C# library for programmatic NVDA control
- Can verify speech output
- Good for regression testing

#### Guidepup
- GitHub: github.com/guidepup/guidepup
- JavaScript library for screen reader automation
- Cross-platform (NVDA, JAWS, VoiceOver)
- Modern API

## Accessibility Testing Best Practices

1. **Test with actual screen reader** - keyboard-only testing misses speech issues
2. **Test all keyboard shortcuts** - verify announcements at each step
3. **Test error states** - what happens when PowerPoint closes mid-use?
4. **Test with real content** - use presentations with actual comments
5. **Document expected behavior** - write down what NVDA should say
