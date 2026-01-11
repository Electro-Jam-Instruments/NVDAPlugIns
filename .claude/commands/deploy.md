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
- The workflow trigger pattern is `*-v[0-9]+.[0-9]+.[0-9]+*` which REQUIRES the plugin name prefix

**Common Mistakes to Avoid:**
- Using `v0.0.X-beta` instead of `powerpoint-comments-v0.0.X-beta` (builds never triggered)
- Forgetting to update manifest.ini version (installed addon doesn't change)
- Not verifying the build actually ran after pushing tag

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
- Download URLs:
  - Latest beta: `https://electro-jam-instruments.github.io/NVDAPlugIns/downloads/powerpoint-comments-latest-beta.nvda-addon`
  - Specific version: `https://electro-jam-instruments.github.io/NVDAPlugIns/downloads/powerpoint-comments-{VERSION}.nvda-addon`
- Next step: Download, install, restart NVDA, test
