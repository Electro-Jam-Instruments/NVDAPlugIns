# Deploy NVDA Addon

Release a new version of one of the add-ons in this repository by pushing a version tag.
The tag triggers `.github/workflows/release.yml`, which builds the add-on, creates a GitHub
Release and publishes the file to GitHub Pages.

## Arguments
- $ARGUMENTS: `{plugin} {version} [stable] [description]`
  - `{plugin}` - the add-on folder name: `powerpoint-comments`, `spatial-typing-feedback` or `windows-dictation-silence`
  - `{version}` - `X.Y.Z`, e.g. `0.1.2`
  - `stable` - optional. Without it the release is a beta (pre-release)
  - `{description}` - optional, used in the tag message

Example: `/deploy spatial-typing-feedback 0.1.2 Fix crash on voice change`

## Instructions

1. Parse the arguments: $ARGUMENTS
   - If the plugin is missing or is not one of the three folders above, stop and ask.
   - If the version is missing, read the current `addon_version` from `{plugin}/buildVars.py` and ask.
2. Run the checklist below. Stop and report if any step fails.

### Pre-flight Checks
- Run `git status` - the working directory must be clean and on `main`
- Run `git log -1 --oneline` to confirm the latest commit
- Read `{plugin}/buildVars.py` - note the current `addon_version`
  - This is the only place the version lives. `addon/manifest.ini` is generated from
    `manifest.ini.tpl` by scons at build time and is not in git.
- Read `{plugin}/CHANGELOG.md` - check there is an entry for the version (see below)
- Run `python tools/check_addons.py` - the same static checks as the `Tests` workflow,
  including "buildVars version is mentioned in CHANGELOG.md"
- If deploying `spatial-typing-feedback`: run `python -m pytest spatial-typing-feedback/tests`
- Build locally so a broken build is caught before a tag goes out:
  `(cd {plugin} && python -m SCons -c && python -m SCons)` (or `scons` if it is on PATH).
  It must end with `Generating Addon {addon_name}-{VERSION}.nvda-addon` and the file must
  exist. The `.nvda-addon` and generated `addon/manifest.ini` are gitignored, so the
  working tree stays clean.
- Check `{plugin}/CHANGELOG.md` has a `## [X.Y.Z-beta]` header (or `## [X.Y.Z]` for stable)
  holding this release's entries. If it reads `- not yet tagged`, replace that with today's
  date (YYYY-MM-DD). Commit it with the version update below, or on its own
  (`git commit -m "Date {plugin} X.Y.Z changelog"`) if there is no version update, and
  push it before tagging.

### Version Update (if the requested version differs from buildVars)
- Update `addon_version` in `{plugin}/buildVars.py` to `X.Y.Z` (no `-beta` suffix - beta is
  set by the tag, not by buildVars)
- Update `{plugin}/CHANGELOG.md`:
  - The release notes are taken from the section headed `## [X.Y.Z-beta]` for a beta tag,
    or `## [X.Y.Z]` for a stable tag. Move the `[Unreleased]` entries under that header.
  - If the header is missing, the release still builds, but its notes just say
    "See CHANGELOG.md for details."
- Re-run `python tools/check_addons.py`
- Commit only those two files: `git add {plugin}/buildVars.py {plugin}/CHANGELOG.md`, then
  `git commit -m "Bump {plugin} to X.Y.Z"`
- Push: `git push`
- Confirm the `Tests` workflow passes on that push before tagging:
  `gh run list --workflow=tests.yml --branch main --limit 1`
  (tag pushes do not run `Tests`, only `release.yml`)

### Create and Push Tag
**CRITICAL: Use the correct format!**
- Beta: `{plugin}-vX.Y.Z-beta`
- Stable: `{plugin}-vX.Y.Z`
- NOT `vX.Y.Z-beta` - without the plugin prefix the workflow cannot tell which add-on to
  build. The trigger patterns are `*-v[0-9]+.[0-9]+.[0-9]+` and `*-v[0-9]+.[0-9]+.[0-9]+-beta`.
- `X.Y.Z` in the tag MUST equal `addon_version` in `{plugin}/buildVars.py`, or the
  "Validate buildVars version" step fails the build.

**Common Mistakes to Avoid:**
- Tagging without the plugin prefix (build never triggers)
- Tagging before the buildVars bump is pushed (version validation fails)
- Adding `-beta` to `addon_version` in buildVars (validation only compares `X.Y.Z`)
- Not verifying the build actually ran after pushing the tag

Commands (beta):
```bash
git tag -a {plugin}-v{VERSION}-beta -m "v{VERSION}-beta: {DESCRIPTION}"
git push origin {plugin}-v{VERSION}-beta
```

Commands (stable):
```bash
git tag -a {plugin}-v{VERSION} -m "v{VERSION}: {DESCRIPTION}"
git push origin {plugin}-v{VERSION}
```

### Verify Build
- Run: `gh run list --workflow=release.yml --limit 1` (the workflow is named "Build NVDA Addon")
- Confirm the run is for the new tag and shows "in_progress" or "completed success"
- If no run appears, the tag format is WRONG - check and fix
- If it failed, `gh run view {RUN_ID} --log-failed` - the usual cause is a tag/buildVars version mismatch
- Check the release exists: `gh release view {TAG}` (betas are marked pre-release)

### Post-Deploy Verification
The workflow pushes to the `gh-pages` branch; GitHub Pages then takes a minute or two to update.
- Beta: `curl -sI "https://electro-jam-instruments.github.io/NVDAPlugIns/downloads/{plugin}-latest-beta.nvda-addon" | grep -i Last-Modified`
- Stable: `curl -sI "https://electro-jam-instruments.github.io/NVDAPlugIns/downloads/{plugin}-latest.nvda-addon" | grep -i Last-Modified`
- Confirm the timestamp is recent
- Optional: `curl -s "https://electro-jam-instruments.github.io/NVDAPlugIns/api/{plugin}.json"` shows the published version and tag

### Report to User
Provide summary:
- Plugin, version and tag deployed (beta or stable)
- Build status
- Download URLs:
  - Latest beta: `https://electro-jam-instruments.github.io/NVDAPlugIns/downloads/{plugin}-latest-beta.nvda-addon`
  - Latest stable: `https://electro-jam-instruments.github.io/NVDAPlugIns/downloads/{plugin}-latest.nvda-addon`
  - Specific version: `https://electro-jam-instruments.github.io/NVDAPlugIns/downloads/{addon_name}-{VERSION}.nvda-addon`
    (`{addon_name}` is from buildVars, e.g. `powerPointComments`, `spatialTypingFeedback`,
    `windowsDictationSilence` - not the folder name)
  - GitHub Release: `https://github.com/Electro-Jam-Instruments/NVDAPlugIns/releases/tag/{TAG}`
- Next step: Download, install, restart NVDA, test

Note: the download page (`index.html`) is hardcoded in `release.yml`. A new add-on beyond
these three builds and publishes, but gets no card on the page until that HTML is edited.
