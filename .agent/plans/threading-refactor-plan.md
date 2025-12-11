# Plan: Move COM Work to Dedicated Background Thread

## Overview

Refactor the PowerPoint Comments addon to use a dedicated background thread for all COM operations, following NVDA maintainer recommendations.

## Current State (v0.0.13)

```
event_appModule_gainFocus()
    └── core.callLater(100ms)
            └── _deferred_initialization()
                    └── comHelper.getActiveObject()
                    └── _has_active_presentation()
                    └── _ensure_normal_view()
```

**Problems with current approach:**
- `core.callLater()` is not recommended for continuous work
- Each focus event creates a new deferred call
- No lifecycle management (no cleanup on terminate)
- Future Phase 2 needs continuous slide monitoring - callLater won't scale

## Proposed Architecture

```
__init__()
    └── Start background thread (PowerPointMonitorThread)

Background Thread (runs continuously):
    └── CoInitializeEx(COINIT_APARTMENTTHREADED)  # STA for Office COM
    └── Loop while not stopped:
            └── Check work queue for tasks
            └── Execute COM operations
            └── Queue UI messages to main thread via queueHandler
            └── Wait with timeout (check stop event)
    └── CoUninitialize()

event_appModule_gainFocus()
    └── Queue "initialize" task to background thread

terminate()
    └── Signal thread to stop
    └── Join thread with timeout
```

## Key Components

### 1. PowerPointWorker Thread Class

```python
import threading
from comtypes import CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED
from queueHandler import queueFunction, eventQueue

class PowerPointWorker:
    """Background thread for PowerPoint COM operations."""

    def __init__(self, app_module):
        self._app_module = app_module
        self._stop_event = threading.Event()
        self._work_queue = queue.Queue()
        self._thread = None
        self._ppt_app = None

    def start(self):
        """Start the background thread."""
        self._thread = threading.Thread(
            target=self._run,
            name="PowerPointCommentWorker",
            daemon=False  # Non-daemon for clean shutdown
        )
        self._thread.start()

    def stop(self, timeout=5):
        """Stop the thread gracefully."""
        self._stop_event.set()
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=timeout)

    def queue_task(self, task_name, *args):
        """Queue a task for the background thread."""
        self._work_queue.put((task_name, args))

    def _run(self):
        """Main thread loop - runs in background."""
        # Initialize COM in STA mode (required for Office)
        CoInitializeEx(COINIT_APARTMENTTHREADED)
        log.info("PowerPoint worker thread started (COM initialized)")

        try:
            while not self._stop_event.is_set():
                try:
                    # Check for work with timeout
                    task_name, args = self._work_queue.get(timeout=0.5)
                    self._execute_task(task_name, args)
                except queue.Empty:
                    # No work, continue loop
                    pass
                except Exception as e:
                    log.error(f"Worker thread error: {e}")
        finally:
            # Always clean up COM
            self._ppt_app = None
            CoUninitialize()
            log.info("PowerPoint worker thread stopped (COM uninitialized)")

    def _execute_task(self, task_name, args):
        """Execute a queued task."""
        if task_name == "initialize":
            self._task_initialize()
        elif task_name == "check_slide":
            self._task_check_slide()
        # Add more tasks as needed

    def _task_initialize(self):
        """Connect to PowerPoint and check presentation."""
        try:
            self._ppt_app = comHelper.getActiveObject(
                "PowerPoint.Application",
                dynamic=True
            )
            log.info("Worker: Connected to PowerPoint COM")

            if self._has_active_presentation():
                self._ensure_normal_view()
        except Exception as e:
            log.error(f"Worker: Initialize failed - {e}")
            self._ppt_app = None

    def _announce(self, message):
        """Safely announce message on main thread."""
        queueFunction(eventQueue, ui.message, message)
```

### 2. Updated AppModule

```python
class AppModule(AppModule):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._worker = PowerPointWorker(self)
        self._worker.start()
        log.info(f"PowerPoint Comments AppModule instantiated (v{ADDON_VERSION})")

    def event_appModule_gainFocus(self):
        """Queue initialization to background thread."""
        log.info("App gained focus - queuing initialization")
        self._worker.queue_task("initialize")

    def terminate(self):
        """Clean up background thread."""
        log.info("Terminating - stopping worker thread")
        self._worker.stop(timeout=5)
        super().terminate()
```

## Threading Safety Rules

### 1. COM Apartment Threading
- Must call `CoInitializeEx(COINIT_APARTMENTTHREADED)` at thread start
- Must call `CoUninitialize()` in finally block
- PowerPoint/Office requires STA (Single-Threaded Apartment)

### 2. UI Calls from Thread
- NEVER call `ui.message()` directly from background thread
- Always use: `queueFunction(eventQueue, ui.message, "text")`
- This queues the call to NVDA's main thread

### 3. Thread Lifecycle
- Use `threading.Event()` for stop signaling
- Non-daemon thread for clean shutdown
- Always join with timeout in terminate()
- Handle exceptions to prevent thread death

### 4. Avoid Deadlocks
- Don't hold locks during COM operations
- Don't cache COM objects longer than necessary
- Create fresh COM references for each operation

## Implementation Steps

### Step 1: Add Threading Infrastructure
- Add imports: `threading`, `queue`, `CoInitializeEx`, `CoUninitialize`
- Create PowerPointWorker class
- Add work queue mechanism

### Step 2: Refactor AppModule
- Initialize worker in `__init__`
- Change `event_appModule_gainFocus` to queue task
- Add `terminate()` method for cleanup

### Step 3: Move COM Code to Worker
- Move `_connect_to_powerpoint()` logic to worker
- Move `_ensure_normal_view()` logic to worker
- Replace `ui.message()` with queued calls

### Step 4: Test
- Verify thread starts on PowerPoint launch
- Verify COM connection works from thread
- Verify UI announcements work
- Verify clean shutdown on PowerPoint close
- Verify no speech blocking

## Future Benefits

This architecture enables Phase 2 features:
- Continuous slide change monitoring (thread can poll)
- Comment status announcements without blocking
- @mention detection in background
- Proper cleanup when PowerPoint closes

## Version

This will be v0.0.14.

## Risks and Mitigations

| Risk | Mitigation |
|------|------------|
| Thread doesn't start | Log at start, verify in NVDA log |
| COM fails in thread | Use comHelper, wrap in try/except |
| UI calls crash | Always use queueFunction |
| Thread won't stop | Join with timeout, log warning |
| Deadlock | Don't hold locks, fresh COM refs |

## Decision Required

Should the worker thread:
A) Poll continuously for slide changes (Phase 2 ready)
B) Only respond to queued tasks (simpler, add polling later)

Recommendation: **Option B** - Start simple, add polling in Phase 2.
