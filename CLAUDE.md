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

## AppModule Inheritance Pattern

CRITICAL: Use correct inheritance. See `.agent/experts/nvda-plugins/decisions.md` Decision #6.

```python
# CORRECT
from nvdaBuiltin.appModules.powerpnt import AppModule as BuiltinPowerPointAppModule
class AppModule(BuiltinPowerPointAppModule):
    pass

# WRONG - addon loads but doesn't extend built-in
from nvdaBuiltin.appModules.powerpnt import *
class AppModule(appModuleHandler.AppModule):
    pass
```

## Current Version

As of December 2025: v0.0.8
