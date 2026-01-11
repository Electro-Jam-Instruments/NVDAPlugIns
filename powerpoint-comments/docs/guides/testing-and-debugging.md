# Testing and Debugging Guide

How to test and debug the NVDA PowerPoint Comments addon.

## NVDA Log Location

```
%TEMP%\nvda.log
```

Or view in NVDA: Menu > Tools > View Log (NVDA+F1)

## Enable Debug Logging

In NVDA Settings > Advanced:
- Set log level to "Debug"
- Enable "Developer Scratchpad" for fast iteration

## Testing Workflow

### Quick Iteration (Scratchpad)

1. Copy `powerpnt.py` to `%APPDATA%\nvda\scratchpad\appModules\`
2. Reload plugins: **NVDA+Ctrl+F3**
3. Test in PowerPoint
4. Check NVDA log for errors
5. Repeat

### Full Addon Testing

1. Build: `scons` in addon directory
2. Install: Double-click `.nvda-addon` file
3. Restart NVDA
4. Test in PowerPoint

## Adding Debug Logging

```python
import logging
log = logging.getLogger(__name__)

# At module load
log.info(f"PowerPoint Comments addon: Module loading (v{ADDON_VERSION})")

# In methods
log.debug(f"Event fired: {event_name}")
log.info(f"Slide changed to {slide_index}")
log.error(f"COM error: {e}")
```

## Common Issues and Diagnostics

### Module Not Loading

**Symptom:** No log entries from your addon

**Check:**
1. Is manifest.ini valid? (quoting rules!)
2. Is the inheritance pattern correct? (see nvda-addon-development.md)
3. Any import errors? Check NVDA log at startup

### COM Access Failing

**Symptom:** `WinError -2147221021 Operation unavailable`

**Fix:** Use `comHelper.getActiveObject()` not direct `GetActiveObject()`

### Events Not Firing

**Symptom:** COM event handlers never called

**Check:**
1. Is the event sink connected? Log in connection code
2. Is message pump running? Events need Windows message processing
3. Is the connection reference kept alive? Don't let it get garbage collected

### Speech Blocked

**Symptom:** NVDA stops speaking during your code

**Fix:** Don't do heavy work in event handlers. Delegate to worker thread:
```python
def event_appModule_gainFocus(self):
    # Non-blocking - just signals worker thread
    if self._worker:
        self._worker.request_initialize()
```

## Test Scenarios

### Edit Mode
1. Open PowerPoint with a presentation
2. Navigate between slides (PageUp/PageDown)
3. Verify "has X comments" / "has notes" announcements
4. Test Ctrl+Alt+N for notes reading

### Slideshow Mode
1. Start slideshow (F5)
2. Navigate slides (arrows, PageDown)
3. Verify only title + prefix announced (not full content)
4. Exit slideshow (Escape)
5. Verify edit mode announcements resume

### Edge Cases
- Presentation with no slides
- Slide with no comments/notes
- Multiple presentations open
- Switching between PowerPoint windows

## Log Analysis Tips

Search for these patterns in NVDA log:

```
# Addon loading
"PowerPoint Comments addon"

# Slide changes
"WindowSelectionChange"
"SlideShowNextSlide"

# Errors
"error"
"exception"
"failed"
```

## Remote Debugging

If testing on another machine:
1. Copy NVDA log file after reproducing issue
2. Check timestamps to correlate with user actions
3. Look for stack traces after exceptions
