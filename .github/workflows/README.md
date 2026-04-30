# GitHub Actions Workflows

This directory contains all CI/CD and automation workflows for RedSalamander.

---

## Core Build & CI

### `ci.yml` — CI
Runs on every push and pull request targeting `main`.

**Jobs:**
- **migration-guards** — Runs three static audit scripts (`Verify-NoSubclassManager.ps1`, `Audit-ComctlReportSurfaces.ps1`, `Audit-VisibleNativeSurfaces.ps1`) to enforce migration compliance.
- **format** — Installs LLVM/clang-format, runs `format-all.ps1`, and auto-commits any formatting fixes back to the branch.
- **build** — Calls `build-reusable.yml` to build the full solution (Release, x64) without uploading artifacts.
- **selftest-build** — Calls `build-reusable.yml` in Debug/x64 mode and uploads the build output as an artifact.
- **selftest** — Downloads the Debug build, stages Themes, runs `RedSalamander.exe --compare-selftest --fileops-selftest`, then runs `ViewerPETests.exe`, `DxUiTests.exe`, and `ViewerSqliteTests.exe`. Uploads selftest artifacts from `%LOCALAPPDATA%\RedSalamander\SelfTest\` on success or failure.

---

### `build-reusable.yml` — Build (reusable)
Reusable workflow called by `ci.yml` and `release.yml`. Accepts `configuration`, `platform`, `build_number`, `official_release`, and `upload_build_output` inputs.

**Steps:**
1. Installs Visual Studio 2026 Build Tools + VC++ workload via Chocolatey.
2. Locates MSBuild for VS 2026 (version range `[18.0, 19.0)`) using `vswhere.exe` and adds it to `$PATH`.
3. Caches and bootstraps vcpkg, validates the `builtin-baseline` commit from `vcpkg.json`, and unshallows the vcpkg repo if needed.
4. Installs vcpkg dependencies and enables MSBuild integration.
5. Normalizes the vcpkg install root so MSBuild can find headers and libs regardless of runner layout.
6. Calls `build.ps1` with the appropriate `-Configuration`, `-Platform`, `-BuildNumber`, and `-OfficialRelease` flags. Injects vcpkg include/lib paths via `$env:CL` and `$env:LINK` to handle runners that lack VS vcpkg integration.
7. Optionally uploads the build output as artifact `build-output-<platform>` (retention: 1 day).

---

### `release.yml` — Release
Triggers on `v*` tag pushes or `workflow_dispatch`. Supports optional ARM64 build, MSI, and MSIX packaging.

**Jobs:**
1. **version-info** — Resolves `build_number`, `release_tag`, and `release_version` from the tag or dispatch input.
2. **build** — Matrix build (x64 + optional ARM64) via `build-reusable.yml` in Release mode. Artifacts uploaded as `build-output-<platform>`.
3. **package-portable** — Downloads build output and calls `Installer\zip\build-zip.ps1` to produce a portable ZIP. Artifact: `portable-package-<platform>`.
4. **package-msi** — Conditional (`build_msi=true`). Installs WiX CLI via `winget`, calls `Installer\msi\build-msi.ps1` and `build-msi-symbols.ps1`. Artifact: `msi-package-<platform>`.
5. **package-msix** — Conditional (`build_msix=true`). Installs VS 2026 + UWP workload, generates MSIX assets, builds the `.wapproj`, and optionally signs with `MSIX_SIGNING_CERT`/`MSIX_SIGNING_PASSWORD` secrets. Artifact: `msix-package-<platform>`.
6. **release** — Downloads all package artifacts, generates SHA256 checksums, and publishes a GitHub Release via `softprops/action-gh-release@v2`. Uploads MSIX separately if that job succeeded.

---

### `winget-release.yml` — Publish to Winget
Triggers when a GitHub Release is published, or manually via `workflow_dispatch` (version input required).

**Steps:**
1. Resolves the version from the release tag or manual input.
2. Downloads the MSI and portable ZIP from the GitHub Release.
3. Installs `winget-create` and submits the update to the winget community repository using the `WINGET_TOKEN` secret.

> **Requires:** `WINGET_TOKEN` repository secret with a GitHub PAT that has permission to open PRs on `microsoft/winget-pkgs`.

---

## Squad Agent Workflows

These workflows power the Squad AI team system. They operate on issue/label events and read `.squad/team.md` for team roster and routing configuration.

### `squad-triage.yml` — Squad Triage
Triggers when any issue receives the `squad` label. Reads `.squad/team.md` to determine the team roster and @copilot capability profile (good fit / needs review / not suitable keywords). Posts a triage comment and routes the issue to the appropriate member.

### `squad-issue-assign.yml` — Squad Issue Assign
Triggers when an issue receives a `squad:<member>` label. Looks up the member in `.squad/team.md` and posts an assignment acknowledgment comment. For `squad:copilot`, triggers a Copilot coding assignment.

### `squad-label-enforce.yml` — Squad Label Enforce
Triggers on any issue label event. Enforces mutual exclusivity within managed label namespaces (`go:`, `release:`, `type:`, `priority:`). Automatically removes conflicting labels and posts a comment when a label is replaced. Also auto-applies `release:backlog` when `go:yes` is added to an issue that has no release target label.

### `sync-squad-labels.yml` — Sync Squad Labels
Triggers on changes to `.squad/team.md` or manually. Parses the members table and ensures `squad:<member>` GitHub labels exist for every roster entry, creating them if absent. Also creates/updates the static `go:`, `release:`, `type:`, and `priority:` label sets.

### `squad-heartbeat.yml` — Squad Heartbeat (Ralph)
Runs on a 30-minute cron schedule, on issue close/label events, and on PR close events. Executes `.squad/templates/ralph-triage.js` to detect untriaged issues and apply triage labels and comments automatically.

---

## Unadapted Squad Templates

These workflows were scaffolded by the Squad agent and have **not yet been adapted** for this Windows C++ project. They use `ubuntu-latest` runners and contain TODO placeholders. They are currently inert (all steps `echo` a message and exit 0), but they will trigger on the events listed below if the corresponding branches exist.

| Workflow | Trigger | Notes |
|---|---|---|
| `squad-ci.yml` | PRs to `dev`, `preview`, `main`, `insider`; push to `dev`, `insider` | Replace placeholder with `build.ps1` on a `windows-latest` runner. |
| `squad-release.yml` | Push to `main` | Replace placeholder with release automation. Currently runs alongside `release.yml` on every push to `main` (no-op). |
| `squad-preview.yml` | Push to `preview` branch | Replace placeholder with validation steps. Requires a `preview` branch to be meaningful. |
| `squad-insider-release.yml` | Push to `insider` branch | Replace placeholder with insider/pre-release steps. Requires an `insider` branch. |
| `squad-docs.yml` | Push to `preview` branch (paths: `docs/**`); `workflow_dispatch` | Replace placeholder with documentation build/deploy steps. |
| `squad-promote.yml` | `workflow_dispatch` (dry_run option) | Merges `dev → preview → main`. Currently broken: references `package.json` and `CHANGELOG.md` which do not exist at the repo root. Requires `dev`, `preview`, and `insider` branches. Do not run until adapted. |
