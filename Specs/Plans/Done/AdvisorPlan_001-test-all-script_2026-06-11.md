# Advisor Plan 001 - Document Run-AllTests discoverability

> **NON-NORMATIVE COMPLETE RECORD.** Archived from root `plans/001-test-all-script.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** COMPLETE
- **Archived:** 2026-08-25
- **Original path:** `plans/001-test-all-script.md`
- **Live owner:** none (implemented; see README.md, CLAUDE.md, AGENTS.md, and Tools/Run-AllTests.ps1)
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 001: Make `Tools\Run-AllTests.ps1` discoverable and document its CI-parity status

> **Frozen historical instructions (do not execute)**: This plan is archived; do not follow these steps. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. Historical: status-row updates no longer apply for this plan
> in `plans/README.md`.
>
> **Drift check (run first)**: `git diff --stat 717ee44d0..HEAD -- Tools/Run-AllTests.ps1 Tools/TestRunPlan.ps1 .github/workflows/ci.yml README.md CLAUDE.md AGENTS.md`
> If any of those changed since this plan was written, re-read them and compare
> against the "Current state" excerpts below before proceeding; on a material
> mismatch (esp. the `-Suite Full` plan in `TestRunPlan.ps1` or the CI step
> list), treat it as a STOP condition.

> **History note (read once, then ignore)**: An earlier version of this plan
> proposed *creating* a new `test-all.ps1` because "no one-command verification
> exists" and "the orphaned `RedConfigureTests`/`FileSystemCurlTests` exes run
> nowhere." **Both premises were false.** `Tools\Run-AllTests.ps1` already is the
> one-command orchestrator, and `-Suite Full` already runs both of those exes
> (plus more than CI does). That obsolete deliverable is retired. This plan is
> now a focused docs/discoverability change. Do NOT create `test-all.ps1`.

## Status

- **Priority**: P2
- **Effort**: S
- **Risk**: LOW (documentation only)
- **Depends on**: none. (Note: this plan is NOT a prerequisite for any other plan — the verification capability it documents already exists.)
- **Category**: dx / docs
- **Planned at**: commit `717ee44d0`, 2026-06-12

## Why this matters

`Tools\Run-AllTests.ps1` already does everything a "run the whole test suite locally" command should: it builds, runs the three in-process selftest suites, and with `-Suite Full` also runs the standalone test exes, the `PerformanceTests2` CppUnitTest DLL, the `Tools\Tests` Pester suite, and the vcpkg-merge synthetic test — then prints a color summary with pass/fail/skip counts and writes an aggregate `run-all-tests-results.json`. The problem is **nobody can find it**: `README.md`, `CLAUDE.md`, `AGENTS.md`, and `docs/` contain zero mentions of it. The README's "Self-tests (Debug only)" section documents only raw `RedSalamander.exe --selftest` invocations, so a developer or coding agent reproducing CI either reverse-engineers `ci.yml` by hand or under-tests. In an agent-driven repo this single missing pointer has outsized cost. Secondarily, the `-Suite Full` plan and the CI selftest job have quietly diverged in both directions; recording how they relate keeps them from rotting apart.

## Current state

### The orchestrator (already exists, do not rewrite)

- `Tools\Run-AllTests.ps1` — parameters at `Tools\Run-AllTests.ps1:60-79`:
  `-Suite` (ValidateSet `All`/`Compare`/`Commands`/`FileOps`/`Full`, default `All`), `-SkipBuild`, `-FailFast`, `-TimeoutMultiplier <double>`, `-CaseFilter <string>` (passed as `--selftest-case=`, prefix match when it ends in `_`), `-Platform` (`x64`/`ARM64`), `-Configuration` (default `Debug`), `-ExePath`.
  - Builds via `build.ps1` unless `-SkipBuild` (`:375-408`); runs selftest suites with `Start-Process -Wait -PassThru` and parses each suite's `results.json` plus a case-coverage check (`:437-571`); runs standalone entries (`:573-588`); writes `run-all-tests-results.json` and a summary (`:595-750`).
- The run plan is built by `Get-RSTestRunPlan` in `Tools\TestRunPlan.ps1:343-447`. **What `-Suite Full` runs** (`:405-444`):
  - Self-tests: `CompareDirectories`, `Commands`, `FileOperations`.
  - Executables: `DxUiTests`, `FileSystemCurlTests`, `ViewerPETests`, `ViewerSqliteTests`, `MonitorTest`, `LocalizationTests`, `RedConfigureTests`.
  - `RedSalamanderMonitorEtwLatency` — `RedSalamanderMonitor.exe --chrome-selftest --perf --monitor-etw-burst-mode=latency …`.
  - CppUnitTest: `PerformanceTests2.dll` (located via `Resolve-VSTestConsole`, `Run-AllTests.ps1:146-184`).
  - Pester: `Tools\Tests` (`Invoke-Pester -Path … -PassThru`, **no** `-ExcludeTag`).
  - PowerShellScript: `Tests\vcpkg-merge-synthetic-test.ps1`.
  - For `-Suite Full`, the build step omits `-ProjectName` so the **whole solution** builds, and sets env `RSBuildEnableTests=true` (`TestRunPlan.ps1:125-145`). So `Run-AllTests.ps1 -Suite Full` (without `-SkipBuild`) is self-contained — it builds the test projects it then runs.

### What CI runs (reference only — do NOT edit ci.yml in this plan)

- `.github/workflows/ci.yml` `selftest` job (`:87-196`, steps `:110-182`): compare+fileops selftests, commands selftest, `ViewerPETests.exe` **×3** (bare + two explicit named cases `TestViewerTextFindPromptUsesDxUiHostAndClosesCleanly` and `TestViewerTextGotoPromptUsesDxUiHostAndClosesCleanly`, `:126-128`), `DxUiTests.exe`+`ViewerSqliteTests.exe`, `MonitorTest.exe`+`LocalizationTests.exe`, `PerformanceTests2.dll` via vstest, `Invoke-Pester .\Tools\Tests -ExcludeTag RequiresBuildToolchain` under **`shell: powershell`** (Windows PowerShell 5.1, `:170-177`), and `Tests\vcpkg-merge-synthetic-test.ps1`.

### How `-Suite Full` and CI relate (the parity note to write)

- **Full covers MORE than CI**: `RedConfigureTests` and `FileSystemCurlTests` (CI runs neither — plan 011 owns adding them to CI), and the `RedSalamanderMonitorEtwLatency` chrome/ETW selftest.
- **CI covers things Full does not mirror exactly**:
  1. The two explicit named `ViewerPETests` cases. Full runs `ViewerPETests.exe` bare; whether the bare run already executes those two cases by default is **unverified** — an executor closing this gap must check (e.g. `ViewerPETests.exe --list` or reading the test's default registration). Treat as a documentation caveat, not a blocker for the docs deliverable.
  2. Pester tag/host: Full runs all `Tools\Tests` in `pwsh` with no tag filter; CI excludes `RequiresBuildToolchain` and runs under Windows PowerShell 5.1. Full is the stricter superset on tags; the PS-host difference could in principle mask a 5.1-only failure.

### Discoverability (the gap this plan closes)

- `Select-String -Path README.md, CLAUDE.md, AGENTS.md -Pattern 'Run-AllTests'` → **no matches** (verified at plan time). `docs\*.md` → none either.
- README "Self-tests (Debug only)" section begins at `README.md:160`; "### Run self-tests" at `:168`; it shows only raw `…\RedSalamander.exe --selftest …` invocations (`:177`, `:185`). This is where the pointer belongs.

## Commands you will need

| Purpose | Command | Expected on success |
|---------|---------|---------------------|
| Confirm script exists | `Test-Path .\Tools\Run-AllTests.ps1` | `True` |
| Help renders | `Get-Help .\Tools\Run-AllTests.ps1 -Full` | renders, no error |
| Discoverability gate (pre) | `Select-String -Path README.md,CLAUDE.md,AGENTS.md -Pattern 'Run-AllTests'` | no matches before, ≥1 each after |
| Optional smoke (cheap suite) | `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_pane_batchRename_window_invokes_success_callback` | exit 0, 1 PASS |

## Scope

**In scope** (the only files you may modify):
- `README.md` — add a "Run all tests with one command" subsection under "Self-tests (Debug only)".
- `CLAUDE.md` — add a one-line pointer (e.g. under Build Commands) to the canonical local verification command.
- `AGENTS.md` — add a one-line pointer so agents know the "is the tree green?" command.
- (Optional, only if it reads cleanly) a short comment block in `Tools\Run-AllTests.ps1`'s `.DESCRIPTION` or `Tools\TestRunPlan.ps1` `Get-RSTestRunPlan` recording the CI-parity relationship. This is the lowest-risk home for the parity note since it sits next to the list it describes.

**Out of scope** (do NOT touch):
- `.github/workflows/ci.yml` and `build-reusable.yml` — CI changes (running the orphan exes, ARM64 gate, or making CI *call* `Run-AllTests.ps1`) belong to **plan 011** / a future convergence plan. Editing CI here causes rebase conflicts with 011.
- The *logic* of `Run-AllTests.ps1` / `TestRunPlan.ps1` — no behavior changes. (A doc-comment is the only allowed touch, and only if optional step 4 is done.)
- Any test source or `build.ps1`.
- Reconciling the divergences (adding the named ViewerPETests cases or the Pester tag filter to the Full plan) — that is a behavior change to the run plan; document the divergence, do not "fix" it here.

## Git workflow

- Branch: `advisor/001-runalltests-discoverability`
- Commit style: short imperative subject matching repo history (e.g. "Document Run-AllTests.ps1 local verification entry point").
- Do NOT push or open a PR unless the operator instructed it.

## Steps

### Step 1: Verify current state

```powershell
Test-Path .\Tools\Run-AllTests.ps1                       # True
Get-Help .\Tools\Run-AllTests.ps1 -Full | Out-Null       # no error
Select-String -Path README.md,CLAUDE.md,AGENTS.md -Pattern 'Run-AllTests'   # expect: no matches
```

**Verify**: script exists, help renders, and the three docs do not yet mention it. If any already mention it, the discoverability work is partly done — reconcile rather than duplicate, and note it in your report.

### Step 2: README — add the one-command subsection

Under the "Self-tests (Debug only)" section (after the existing raw-invocation examples, around `README.md:160-218`), add a subsection titled e.g. **"Run all tests with one command"**. Content (adapt prose to the README's voice; keep these facts exact):

- The canonical local runner is `Tools\Run-AllTests.ps1`. It builds (unless `-SkipBuild`), runs the selected suites, and prints a pass/fail/skip summary; exit code 0 = all green.
- Show these examples:
  ```powershell
  # Everything CI runs (and a bit more) — builds the whole solution with tests enabled:
  .\Tools\Run-AllTests.ps1 -Suite Full

  # Just the three in-process selftest suites (default), reusing an existing Debug build:
  .\Tools\Run-AllTests.ps1 -SkipBuild

  # One suite, one case, on a slow machine:
  .\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_pane_batchRename_ -TimeoutMultiplier 2.0
  ```
- Note that `-Suite Full` builds the full solution (test projects included) and runs the standalone test exes, `PerformanceTests2`, the `Tools\Tests` Pester suite, and the vcpkg-merge synthetic test in addition to the selftests; results land under `%LOCALAPPDATA%\RedSalamander\SelfTest\last_run\`.
- Note selftests are Debug-only (the runner defaults to Debug).

**Verify**: `Select-String -Path README.md -Pattern 'Run-AllTests'` → ≥1 match.

### Step 3: CLAUDE.md and AGENTS.md — one-liner each

- `CLAUDE.md`: under the "## Build Commands" area, add one line, e.g.: *"Run the full local test suite with `\.\Tools\Run-AllTests.ps1 -Suite Full` (or `-SkipBuild` to reuse an existing Debug build). See README "Self-tests" for details."*
- `AGENTS.md`: add an equivalent one-liner where build/verification guidance lives (near the "## Build" section) so an agent's "is it green?" reflex points at the real command.

**Verify**: `Select-String -Path CLAUDE.md,AGENTS.md -Pattern 'Run-AllTests'` → ≥1 match in each.

### Step 4 (optional, low-risk): record the CI-parity note next to the plan

If it reads cleanly, add a short comment in `Tools\TestRunPlan.ps1` near `Get-RSTestRunPlan`'s `Full` block (`:405`) summarizing the relationship documented in "Current state → How `-Suite Full` and CI relate" above: Full is a superset of CI on the standalone exes (adds RedConfigure/FileSystemCurl/ETW-latency), and the two known divergences (the named ViewerPETests cases; the Pester tag/host difference). This is documentation only — change no code. Skip this step if it would mean restructuring the function.

**Verify**: if done, `git diff Tools/TestRunPlan.ps1` shows only comment lines added.

### Step 5: Final verification

```powershell
Select-String -Path README.md,CLAUDE.md,AGENTS.md -Pattern 'Run-AllTests' | Select-Object Path -Unique
git status --short
```

**Verify**: all three docs match; `git status` shows only in-scope files modified.

## Test plan

Documentation change — no automated tests. Verification is the grep gates in the Done criteria plus the optional cheap smoke (`-Suite Commands -SkipBuild -CaseFilter …`) to confirm the documented command actually works as written before you ship docs that tell others to run it.

## Done criteria

Machine-checkable. ALL must hold:

- (archived, not live)  `Select-String -Path README.md -Pattern 'Run-AllTests'` → ≥1 match, in a subsection that shows `-Suite Full` and `-SkipBuild`
- (archived, not live)  `Select-String -Path CLAUDE.md -Pattern 'Run-AllTests'` → ≥1 match
- (archived, not live)  `Select-String -Path AGENTS.md -Pattern 'Run-AllTests'` → ≥1 match
- (archived, not live)  The documented smoke command runs: `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_pane_batchRename_window_invokes_success_callback` exits 0 (requires an existing Debug build; if none, build first or note it skipped)
- (archived, not live)  No files outside the in-scope list modified (`git status`); in particular `ci.yml` untouched
- (archived, not live)  `plans/README.md` status row for 001 updated

## STOP conditions

Stop and report back (do not improvise) if:

- `Tools\Run-AllTests.ps1` or `Get-RSTestRunPlan` in `Tools\TestRunPlan.ps1` no longer matches the "Current state" excerpts (the script was refactored since `717ee44d0`) — re-derive the real `-Suite Full` inventory before documenting it; documenting a stale command set is exactly the failure this rewrite corrects.
- The optional smoke command fails — the script may be broken at HEAD; report rather than shipping docs that point at a broken command.
- You find the docs already point at `Run-AllTests.ps1` (someone did this already) — reconcile/improve rather than duplicate, and say so.

## Maintenance notes

- **The real long-term fix is convergence**: CI's `selftest` job and `Get-RSTestRunPlan`'s `Full` branch are two hand-maintained lists of the same suite set and have already drifted. Having CI *call* `Tools\Run-AllTests.ps1 -Suite Full` would collapse them to one list. That is deferred (it needs CI artifact-layout alignment) and is noted in plan 011's maintenance section; it is **not** this plan.
- **Cross-plan stale references (reconciled 2026-06-12)**: plans 003, 004, 005, 006, and 007 originally referenced a `test-all.ps1` that will never exist. They were corrected to point at the real runner `Tools\Run-AllTests.ps1` and, where a plan creates a new test exe (004/006/007), to register it in `Tools\TestRunPlan.ps1` `Get-RSTestRunPlan` (the `Full` block, `:405-412`). If a future plan adds a standalone test exe, follow that same pattern — there is no per-exe `-Suite` value and no `test-all.ps1`.
- Reviewer focus: that the README examples are copy-pasteable and exact (suite names, parameter spellings), and that `ci.yml` was not touched.
