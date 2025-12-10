# Deploy NVDA Addon

Deploy a new version of the powerpoint-comments addon.

## Arguments
- $ARGUMENTS: Version number (e.g., "0.0.9") and optional description

## Instructions

1. Parse the version from arguments: $ARGUMENTS
2. Run the deployment checklist:

### Pre-flight Checks
- Read `powerpoint-comments/addon/manifest.ini` - note current version
- Read `powerpoint-comments/buildVars.py` - verify version matches manifest
- Run `git status` to ensure working directory is clean
- Run `git log -1 --oneline` to confirm latest commit

### Version Update (if version provided differs from current)
- Update `powerpoint-comments/addon/manifest.ini` version field
- Update `powerpoint-comments/buildVars.py` addon_version field
- Commit: `git commit -am "Bump version to X.X.X"`
- Push: `git push`

### Create and Push Tag
**CRITICAL: Use correct format!**
- Tag format MUST be: `powerpoint-comments-vX.X.X-beta`
- NOT: `vX.X.X-beta` (this will NOT trigger the build!)

Commands:
```bash
git tag -a powerpoint-comments-v{VERSION}-beta -m "v{VERSION}-beta: {DESCRIPTION}"
git push origin powerpoint-comments-v{VERSION}-beta
```

### Verify Build
- Run: `gh run list --workflow=build-addon.yml --limit 1`
- Confirm status shows "in_progress" or "completed success"
- If no run appears, the tag format is WRONG - check and fix

### Post-Deploy Verification
- Check GitHub Pages: `curl -sI "https://electro-jam-instruments.github.io/NVDAPlugIns/downloads/powerpoint-comments-latest-beta.nvda-addon" | grep Last-Modified`
- Confirm timestamp is recent (within last few minutes)

### Report to User
Provide summary:
- Version deployed
- Build status
- Download URL: https://electro-jam-instruments.github.io/NVDAPlugIns/
- Next step: Download, install, restart NVDA, test
