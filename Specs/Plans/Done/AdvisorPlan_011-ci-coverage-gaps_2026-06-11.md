# Advisor Plan 011 - Close CI coverage gaps

> **NON-NORMATIVE COMPLETE RECORD.** Archived from root `plans/011-ci-coverage-gaps.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** COMPLETE
- **Archived:** 2026-08-25
- **Original path:** `plans/011-ci-coverage-gaps.md`
- **Live owner:** Specs/Plans/Done/Operation_Lighthouse_WholeRepositoryAuditFindingsAndRemediationRouting_2026-07-10.md (Lighthouse Track D)
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 011: Close CI coverage gaps — run the orphaned test exes, add an ARM64 build gate

> **Retired 2026-07-15:** Superseded and completed by Lighthouse Track D in
> `Specs/Plans/WIP/Operation_Lighthouse_WholeRepositoryAuditFindingsAndRemediationRouting_2026-07-10.md`.
> FileSystemCurlTests and RedConfigureTests were already in Suite CI when Lighthouse began; Track D added
> the remaining PluginContractTests, SettingsSchemaTests, CrashHandlingTests, and Debug ARM64 build-only
> gate. Do not execute this historical plan independently.

> **Frozen historical instructions (do not execute)**: This plan is archived; do not follow these steps. CI changes cannot be
> fully verified locally — the Done criteria distinguish what you verify locally
> (YAML validity, exes pass locally) from what the operator verifies on the
> first CI run. Honor STOP conditions. Historical: status-row updates no longer apply in
> `plans/README.md`.
>
> **Drift check (run first)**: `git diff --stat d6bfccc42..HEAD -- .github/workflows/ci.yml .github/workflows/build-reusable.yml`
> On drift, re-read both workflows before proceeding.

## Status

- **Priority**: P2
- **Effort**: S
- **Risk**: LOW (CI-only; worst case is a red job that gets reverted)
- **Depends on**: none (001 is complementary, not required)
- **Category**: tests / dx
- **Planned at**: commit `d6bfccc42`, 2026-06-11

## Why this matters

Two concrete gaps, both verified by reading the workflows at plan time:
1. `Tests\RedConfigureTests` and `Tests\FileSystemCurlTests` are built by the solution but executed **nowhere** in CI — they are dead weight unless they run.
2. PR CI never builds ARM64 at all (the PR `build` job is Release **x64** only — `build-reusable.yml` defaults), so ARM64-only breakage (alignment, intrinsics, project-file drift) is discovered at release time. An ARM64 *build* gate (no test execution — GitHub-hosted runners can't execute ARM64 anyway) catches that class early.

Explicitly NOT in this plan: running test suites against Release binaries. The three in-process selftest suites are Debug-only by design, and whether the standalone test exes are even built/meaningful in Release is unknown — that investigation is recorded as deferred in plans/README.md.

## Current state

- `.github/workflows/ci.yml` (read in full at plan time):
  - `build:` job → `uses: ./.github/workflows/build-reusable.yml` with only `upload_build_output: false` → inherits defaults `configuration: Release`, `platform: x64` (defaults verified at `build-reusable.yml:6-11`).
  - `selftest-build:` job → Debug/x64, uploads artifact `build-output-x64`.
  - `selftest:` job (needs selftest-build) → downloads artifact into `.build/x64/Debug/`, then runs (lines 110-181): compare+fileops selftests, commands selftest, `ViewerPETests.exe` (×3 invocations), `DxUiTests.exe`, `ViewerSqliteTests.exe`, `MonitorTest.exe`, `LocalizationTests.exe`, PerformanceTests2 via vstest, Pester `Tools\Tests`, `Tests\vcpkg-merge-synthetic-test.ps1`. **No `RedConfigureTests.exe`, no `FileSystemCurlTests.exe`.**
  - Other jobs: `migration-guards`, `format` (auto-commit — plan 012's territory; do not touch it here).
- `.github/workflows/build-reusable.yml:3-28` — `workflow_call` inputs `configuration` (default Release), `platform` (default x64), `build_number`, `official_release`, `upload_build_output`; sets `VCPKG_DEFAULT_TRIPLET` from platform (`arm64-windows` for ARM64 — ARM64 plumbing already exists).
- Timeouts in the selftest job range 10–50 minutes per step; new exe steps should get a sensible `timeout-minutes` (15 matches the Monitor/Localization step).

## Commands you will need

| Purpose | Command | Expected on success |
|---------|---------|---------------------|
| Build Debug x64 | `.\build.ps1` | exit 0 |
| Run the orphaned exes locally | from `.build\x64\Debug`: `.\RedConfigureTests.exe`; `.\FileSystemCurlTests.exe` | exit 0 each |
| YAML sanity | `Get-Content .github\workflows\ci.yml -Raw | ConvertFrom-Yaml` requires a module — if unavailable, use `actionlint` if installed, else careful manual review + the operator's CI run is the gate | parses / no findings |
| ARM64 build smoke (optional, slow) | `.\vcpkg-install.ps1 -Platform ARM64` then `.\build.ps1 -Platform ARM64` | exit 0 |

## Scope

**In scope**:
- `.github/workflows/ci.yml` — add the two test exes to the selftest job; add an `arm64-build` job.

**Out of scope** (do NOT touch):
- The `format` job (plan 012).
- `build-reusable.yml` — its inputs already support ARM64; no changes needed.
- Test sources — if an orphaned exe fails locally at baseline, see STOP conditions.
- Release-configuration test execution (deferred; see plans/README.md).

## Git workflow

- Branch: `advisor/011-ci-coverage`
- Subject: "Run RedConfigure and FileSystemCurl tests in CI; add ARM64 build gate".
- Do NOT push or open a PR unless the operator instructed it (CI verification requires a PR — coordinate with the operator).

## Steps

### Step 1: Local baseline for the orphaned exes

`.\build.ps1`, then run `RedConfigureTests.exe` and `FileSystemCurlTests.exe` from `.build\x64\Debug`.

**Verify**: both exit 0. If either fails, STOP for that exe (see STOP conditions) — do not put a known-red test into CI.

### Step 2: Add them to the selftest job

In `ci.yml`, after the "Run Monitor and Localization Tests" step, add:

```yaml
      - name: Run RedConfigure and FileSystemCurl Tests
        shell: pwsh
        working-directory: .build/x64/Debug
        run: |
          .\RedConfigureTests.exe
          .\FileSystemCurlTests.exe
        timeout-minutes: 15
```

(Indentation and style copied from the adjacent steps at ci.yml:139-145. Note pwsh fails the step on native non-zero exit via `$PSNativeCommandUseErrorActionPreference` only if configured — match EXACTLY how the adjacent multi-exe step handles this; the existing "Run Monitor and Localization Tests" step relies on default behavior, so mirroring it is correct and consistent.)

**Verify**: YAML still parses (see commands table); `git diff` shows only the added step.

### Step 3: Add the ARM64 build gate

Add a job alongside `build:`:

```yaml
  build-arm64:
    uses: ./.github/workflows/build-reusable.yml
    with:
      platform: ARM64
      upload_build_output: false
```

This inherits `configuration: Release`. Trade-off to note in your report: this roughly doubles PR build minutes (vcpkg restore for arm64-windows is the dominant cost; `build-reusable.yml` may have vcpkg caching — check and report whether the cache key includes the triplet so ARM64 gets its own cache rather than thrashing x64's).

**Verify**: YAML parses; job structure mirrors the existing `build:` job exactly.

### Step 4: Optional local ARM64 cross-build smoke

If disk/time allows, `.\vcpkg-install.ps1 -Platform ARM64` + `.\build.ps1 -Platform ARM64` to pre-verify the gate would be green. This is slow (full ARM64 vcpkg restore); skipping is acceptable — say so in the report.

### Step 5: Hand off for CI verification

The real gate is a CI run. Leave the branch ready; tell the operator the first run on a PR verifies: the new test step passes, `build-arm64` completes, and total PR wall-time impact.

## Test plan

- Local: Step 1 exe runs (the new CI step's exact commands).
- CI: first PR run (operator-verified).

## Done criteria

- (archived, not live)  `ci.yml` contains the new test step and the `build-arm64` job; nothing else changed in the file (`git diff` review)
- (archived, not live)  Both exes exit 0 locally at the commit being submitted
- (archived, not live)  YAML parses (or actionlint clean, or careful manual review documented)
- (archived, not live)  Report notes the vcpkg-cache/triplet answer from Step 3 and the expected CI-minutes cost
- (archived, not live)  `plans/README.md` status row updated (status BLOCKED-on-CI-run is acceptable until the operator confirms green)

## STOP conditions

- An orphaned exe fails at baseline: do NOT add it to CI; report the failure output. Add only the passing one (adjust the step accordingly). If both fail, this plan reduces to the ARM64 gate only.
- `FileSystemCurlTests.exe` turns out to need network/secrets (its name suggests it might): run it locally disconnected from any VPN to check; if it needs the network, examine whether it self-skips cleanly; if it hard-fails offline, treat as baseline failure.
- `build-reusable.yml` lacks ARM64-compatible vcpkg caching and the restore exceeds ~40 min in your local smoke — flag to the operator before adding the job; the gate may need `push`-to-master-only triggering instead of every PR (that variant is a one-line `if:` the operator can choose).

## Maintenance notes

- CI's selftest steps and `Tools\Run-AllTests.ps1` (`-Suite Full`, via `Get-RSTestRunPlan` in `Tools\TestRunPlan.ps1`) are two hand-maintained lists of the same suite set and have already drifted (Full runs RedConfigure/FileSystemCurl/ETW-latency that CI doesn't; CI runs two named ViewerPETests cases + a Pester `-ExcludeTag` that Full doesn't). Long-term fix: have CI **call** `Run-AllTests.ps1 -Suite Full` so there is one list (deferred; needs artifact-layout alignment). See plan 001 for the parity detail. NOTE: plan 001 was rewritten — there is no `test-all.ps1`; the runner is `Tools\Run-AllTests.ps1`.
- Reviewer focus: that the new step's failure semantics match the neighboring steps (a silently-green step that ignores exit codes is worse than no step).
