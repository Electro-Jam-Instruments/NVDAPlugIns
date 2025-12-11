# PowerPoint Comments NVDA Addon - Project Instructions

## CRITICAL: Deployment Rules

### Tag Format - MUST FOLLOW EXACTLY
**CORRECT:** `powerpoint-comments-vX.X.X-beta` or `powerpoint-comments-vX.X.X`
**WRONG:** `vX.X.X-beta` or `vX.X.X`

The workflow trigger pattern is `*-v[0-9]+.[0-9]+.[0-9]+*` which REQUIRES the plugin name prefix.

### Version Update Checklist - BEFORE EVERY RELEASE

1. Update version in `powerpoint-comments/addon/manifest.ini`
2. Update version in `powerpoint-comments/buildVars.py` (keep in sync)
3. Commit: `git commit -am "Bump version to X.X.X"`
4. Push: `git push origin PPTCommentReview` (or current branch)
5. Create tag: `git tag -a powerpoint-comments-vX.X.X-beta -m "vX.X.X-beta: Description"`
6. Push tag: `git push origin powerpoint-comments-vX.X.X-beta`
7. Verify build: `gh run list --workflow=build-addon.yml --limit 1`

### Common Mistakes We've Made
- Using `v0.0.X-beta` instead of `powerpoint-comments-v0.0.X-beta` (builds never triggered)
- Forgetting to update manifest.ini version (installed addon doesn't change)
- Not verifying the build actually ran after pushing tag

### Quick Deploy Command Sequence
```bash
# After code changes are committed and pushed:
VERSION="0.0.9"  # Change this to new version
git tag -a powerpoint-comments-v${VERSION}-beta -m "v${VERSION}-beta: Brief description"
git push origin powerpoint-comments-v${VERSION}-beta
gh run watch $(gh run list --workflow=build-addon.yml --limit 1 --json databaseId -q '.[0].databaseId')
```

## Project Context

See RELEASE.md for full release process documentation.
See .agent/agent.md for project overview and expert knowledge locations.

## AppModule Inheritance Pattern - CRITICAL

**USE THE EXACT NVDA DOCUMENTATION PATTERN.** See `.agent/experts/nvda-plugins/decisions.md` Decision #6.

```python
# CORRECT - Exact NVDA documentation pattern (VERIFIED WORKING v0.0.9+)
from nvdaBuiltin.appModules.powerpnt import *

class AppModule(AppModule):  # Inherits from just-imported AppModule
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)  # super() works for __init__
```

**PATTERNS THAT DO NOT WORK:**
```python
# WRONG - Explicit alias (v0.0.4-v0.0.8) - Module does NOT load
from nvdaBuiltin.appModules.powerpnt import AppModule as BuiltinPowerPointAppModule
class AppModule(BuiltinPowerPointAppModule):
    pass

# WRONG - Base class (v0.0.1-v0.0.3) - Loads but loses built-in features
from nvdaBuiltin.appModules.powerpnt import *
class AppModule(appModuleHandler.AppModule):
    pass
```

## Event Handler Pattern - CRITICAL

**`event_appModule_gainFocus` is an OPTIONAL HOOK - parent class does NOT define it.**

```python
# CORRECT - No super() call, defer heavy work
def event_appModule_gainFocus(self):
    # Do NOT call super() - method doesn't exist in parent, will crash
    # Do NOT do heavy work here - blocks NVDA speech
    core.callLater(100, self._deferred_initialization)

# WRONG - Will crash with AttributeError
def event_appModule_gainFocus(self):
    super().event_appModule_gainFocus()  # FAILS - method doesn't exist
```

**Key rules:**
- `super().__init__()` - YES, parent has this method
- `super().event_appModule_gainFocus()` - NO, parent doesn't have this method
- Heavy work in event handlers blocks NVDA speech - always defer with `core.callLater()`

## COM Access Pattern - CRITICAL

**Use `comHelper.getActiveObject()` NOT direct `GetActiveObject()`.**

NVDA runs with UIAccess privileges which prevents direct COM access to lower-privilege processes like PowerPoint. Error: `WinError -2147221021 Operation unavailable`

```python
# CORRECT - Use NVDA's comHelper (VERIFIED WORKING v0.0.13)
import comHelper
ppt_app = comHelper.getActiveObject("PowerPoint.Application", dynamic=True)

# WRONG - Fails with UIAccess privilege error
from comtypes.client import GetActiveObject
ppt_app = GetActiveObject("PowerPoint.Application")  # FAILS
```

## Current Version

As of December 2025: v0.0.14 (Phase 1 Complete - Threading Architecture)
