# Releasing Prime Word (and how updates reach people)

Installed copies of Prime Word check this repository's latest GitHub release once a day (and from
File → About → Check for updates). When it has a higher version, they offer to install it and restart.
Updates install only if they are signed with Prime Word's private key.

## The signing key

- Private key: `%USERPROFILE%\.prime-word\updater.key` (never commit it; it is outside the repository).
- Public key: in `apps/word/src-tauri/tauri.conf.json` → `plugins.updater.pubkey` (built into the app).
- **Back up `%USERPROFILE%\.prime-word\` somewhere safe (a USB drive, a password manager).** If the
  private key is lost, installed copies can never be updated again: everyone would have to download and
  install a new build by hand.

## Making a release

1. Build with a higher version than the last release, and say what is new:

   ```powershell
   powershell -NoProfile -File scripts/build-installer.ps1 -Version 0.1.1 -Notes "Shapes: transparency, adjustment handles. Fixes: ..."
   ```

   This writes `release\v0.1.1\`: `PrimeWord_0.1.1_x64-setup.exe`, its `.sig`, and `latest.json`.

2. On GitHub: **Releases → Draft a new release**, tag `v0.1.1` (exactly `v` + the version), title
   "Prime Word 0.1.1", and attach the three files from `release\v0.1.1\`. Publish it (not as a
   pre-release: the updater reads the *latest* release).

3. That's it: installed copies find it within a day, or at once with File → About → Check for updates.

## Notes

- People with a build from before 0.1.0's updater need to install one updater build by hand once.
- The installer is not code-signed yet, so Windows SmartScreen warns on first install (More info → Run
  anyway). Updates through the app do not show that warning.
- The version lives in `apps/word/src-tauri/tauri.conf.json`; `-Version` sets it. Commit that change
  with the release.
