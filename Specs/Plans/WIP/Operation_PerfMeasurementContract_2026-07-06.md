# Perf Measurement Contract Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development
> (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use
> checkbox (`- [ ]`) syntax for tracking.

**Created:** 2026-07-06
**Status:** WIP (plan authored and linked; short-root contract merged into authoritative testing/archive docs; contract-enforcement tasks not started)
**Operation:** Operation Perf Measurement Contract
**Goal:** Make every perf-sensitive RedSalamander change produce comparable, archived, analyzer-ready
measurement evidence before it can be called complete.
**Architecture:** Keep `Specs/Testing/Testing_PerformanceValidation.md` authoritative, and use this
plan to install the enforcement path around it: reusable Perf Measurement Records, runner metadata,
archive preservation, analyzer quality gates, and closeout checks. Evidence must flow from
deterministic harnesses through `Run-AllTests.ps1` summaries and `Specs/TestRuns/<MachineHash>/<Area>/<RunId>/`
archives so reviewers can compare baseline and candidate runs without relying on terminal memory.
**Tech Stack:** PowerShell 7, Pester, native RedSalamander selftests, JSON/JSONL run summaries,
`Tools/Show-PerfRuns.ps1`, Markdown specs.
**Related:** `Specs/Testing/Testing_PerformanceValidation.md`,
`Specs/Testing/Testing_SelfTests.md`, `Specs/TestRuns/README.md`,
`Tools/Show-PerfRuns.ps1`, `Tools/Run-AllTests.ps1`,
`Specs/Plans/WIP/Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md`.

This plan operationalizes the authoritative performance rules in
`Specs/Testing/Testing_PerformanceValidation.md`. It does not replace that spec.
When this plan finishes, durable rules discovered here must be merged back into the
authoritative testing and subsystem specs before the plan moves to `Specs/Plans/Done/`.

---

## Problem

RedSalamander already has a strong normative perf spec, but WIP plans and closeout notes can still
miss required measurement details:

- the exact user-visible scenario,
- authoritative metric keys,
- deterministic validation command,
- archive path and artifact contents,
- build flavor,
- baseline/candidate comparison,
- analyzer command and sample-quality gate,
- environment matrix,
- blocked reason when evidence cannot be collected.

When any of those are missing, perf work becomes storytelling. The team can accidentally accept
"felt faster," ad hoc terminal output, Debug-only timing, low-sample p95 claims, or machine-specific
numbers that cannot be compared later.

---

## Contract

Every perf-sensitive change, including test-stabilization work that changes waits, timing ceilings,
runner duration, rendering sample collection, search behavior, FileOps behavior, Compare Directories,
plugin I/O, startup, memory retention, or background queueing, must carry a **Perf Measurement Record**.

The record may live in the PR description, closeout note, linked WIP plan section, or archived
`perf-analysis.md`, but it must contain every field below or mark the field `[blocked]` with the exact
missing prerequisite.

```markdown
### Perf Measurement Record

- Scenario:
- Subsystem:
- Change type: feature | optimization | stabilization | regression-fix | instrumentation-only
- User-visible risk protected:
- Metric keys:
- Metric units and sample grain:
- Existing instrumentation reused or new instrumentation added:
- Deterministic validation:
- Build flavor:
- Baseline run:
- Candidate run:
- Same-machine and same-suite comparison:
- Analyzer command:
- Sample-quality result:
- Environment matrix:
- Archived evidence:
- Before/after delta:
- Caveats or blockers:
- Authoritative spec updates required:
```

**Hard rule:** a perf-sensitive change is not complete if the record has no scenario, no metric keys,
no deterministic validation, or no archived evidence/blocked reason.

---

## Required Roots

Perf measurement work has three roots. Do not create an alternate root without updating this plan,
`Specs/Testing/Testing_PerformanceValidation.md`, and `Specs/TestRuns/README.md` in the same change.

- WIP plan root: `Specs/Plans/WIP/Operation_PerfMeasurementContract_2026-07-06.md`.
- Default transient runner root: `REDSALAMANDER_TEST_ROOT=<repoRoot>\.build\TestSandbox`, with one
  `runs\<runId>\` directory per invocation and harness artifacts under `runs\<runId>\artifacts\...`.
- Short NTFS transient root for capability-sensitive local evidence:
  `REDSALAMANDER_TEST_ROOT=C:\RSPerf`. Use this root when the repo workspace is on a non-NTFS,
  redirected, network, or path-length-hostile volume and the scenario needs NTFS semantics or short
  absolute paths. The root must contain only the runner-owned `runs\` directory; every run still uses
  `runs\<runId>\scratch\...` and `runs\<runId>\artifacts\...`.
  Do not use `%LOCALAPPDATA%\Temp\RedSalamander-TestSandbox` for final perf evidence: that profile path
  is long enough to invalidate path-sensitive Compare/search cases. Do not use `C:\RST`; the
  flake-convergence plan documents that name family as historical cleanup debt, so blessing it here
  would blur new evidence with legacy scratch.
  Any `REDSALAMANDER_TEST_ROOT` override must be named in the Perf Measurement Record with the reason,
  filesystem type, and whether the result is final evidence or diagnostic only.
- Durable evidence root: `Specs/TestRuns/<MachineHash>/<Area>/<RunId>/`, including `results.json`,
  `trace.txt`, `perf_metrics.jsonl` when metrics are emitted, and `perf-analysis.md` when the run
  supports a before/after or budget claim.

The transient root is disposable diagnostic material. The durable root is the reviewable evidence
record. A closeout note may cite transient paths only as supporting context; final perf claims must
point to the durable `Specs/TestRuns/` archive or state the exact blocker.

---

## File Map

- Modify: `Specs/Testing/Testing_PerformanceValidation.md` for the durable PR/closeout contract.
- Modify: `Specs/Testing/Testing_SelfTests.md` for runner artifact and archive preservation rules.
- Modify: `Specs/TestRuns/README.md` for archive folder and `perf-analysis.md` conventions.
- Modify: `Specs/Plans/WIP/Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md` for
  the concrete timing/perf records that plan must carry.
- Modify: `Tools/TestRunPlan.ps1` for pure summary-schema and metadata helpers.
- Modify: `Tools/Run-AllTests.ps1` for runtime collection and aggregate-summary emission.
- Test: `Tools/Tests/RunAllTestsPlan.Tests.ps1` for runner summary metadata and archive path rules.
- Test: `Tools/Tests/ShowPerfRuns.Tests.ps1` for analyzer quality-gate behavior.
- Test: `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` for required documentation links and
  no-move-to-Done enforcement.

---

## Measurement Requirements

### 1. Scenario first

The scenario must name the exact user-visible path being protected. Good examples:

- FolderView overlay invalidation while repeated hover/popup repaint traffic is active.
- Local index snapshot reload after service database update.
- FileOps popup lifetime during completion teardown.
- Compare Directories remote connection probe on credentialed-but-unreachable endpoints.

Bad examples:

- "Perf is better."
- "Timing stabilized."
- "Runner faster."

### 2. Metric keys are mandatory

Every record must list metric keys or explain why existing metrics are sufficient. Prefer subsystem
families already used by the repo:

- `folder.frame.*`
- `folder.scroll.*`
- `folder.refresh.*`
- `icons.*`
- `compare.*`
- `find.ui.*`
- `FileOps.*`
- `viewer.diff.*`

Metric rows must state units and sample grain: per frame, per refresh, per query, per operation,
per queue drain, per case, or per run.

If a change only removes an invalid correctness timing ceiling, such as an `elapsed < X ms` assert,
the replacement must still record an advisory metric when the timing remains useful for drift
detection. Do not silently delete the only timing signal.

### 3. Deterministic validation

Use the narrowest deterministic path that exercises the scenario:

- `--commands-selftest` for real UI/window behavior,
- `--compare-selftest` for Compare/search engine behavior,
- `--fileops-selftest` for File Operations,
- standalone harness or Pester only when that is the owner of the behavior.

The record must include the exact command, case filter, build configuration, timeout multiplier,
perf budget path when applicable, and any required environment variables.

For selftest-based Release evidence, use a test-enabled Release build:

```powershell
try {
    $env:RSBuildEnableTests='true'
    .\build.ps1 -ProjectName RedSalamander -Configuration Release
} finally {
    Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue
}
```

### 4. Build flavor

Perf evidence must report the actual build flavor from the JSONL `build` field or the run summary.
Release claims require test-enabled Release evidence unless the authoritative spec explicitly says
Debug evidence is acceptable for that scenario.

Debug evidence is useful for diagnostics and flake reproduction, but it is not sufficient for a
final throughput, latency, p95/p99, or budget claim.

### 5. Archive every meaningful run

Meaningful evidence must be archived under:

```text
Specs/TestRuns/<MachineHash>/<Area>/<RunId>/
```

Minimum artifact set:

- `results.json`
- `trace.txt`
- `perf_metrics.jsonl` when metrics are emitted
- `run-all-tests-results.json` or harness summary when the runner created one
- `perf-analysis.md` with the Perf Measurement Record and analyzer output

If the archive path cannot be created, the plan or closeout must say exactly why and who owns the
follow-up. Unarchived terminal output is not durable evidence.

### 6. Baseline and candidate

Optimization and stabilization claims must compare:

- baseline run,
- candidate run,
- same-machine status,
- same-suite status,
- changed metric values,
- caveats.

If same-machine evidence is unavailable, the result is directional only. Do not promote it to a hard
budget or final improvement claim.

### 7. Analyzer and quality gate

For percentile or frame-time claims, the required first-pass analyzer is:

```powershell
.\Tools\Show-PerfRuns.ps1 <archive-or-jsonl-path> -FailOnQuality
```

Use `-FolderViewPreset` for FolderView frame, scale, cold, slow, scroll, icon, thumbnail, and refresh
families. A p95 claim needs at least 200 samples for that metric unless the authoritative budget
declares a smaller `minimumSamples`. A p99 claim needs at least 1000 samples.

Automation that gates p95/p99 must fail on insufficient sample quality. Low-sample p95 output may be
shown as diagnostic context, but it must not be cited as evidence.

### 8. Environment matrix

Every archived perf run that supports a claim must record:

- machine hash,
- build flavor,
- CPU/load condition,
- active DPI,
- display refresh rate,
- display scale percent,
- local console vs RDP,
- WARP availability and whether WARP was executed,
- adapter name,
- driver-version availability,
- high-DPI run status,
- OS version,
- timeout multiplier,
- relevant feature flags and environment variables.

For CPU-loaded timing robustness, define the load generator, duration, target CPU pressure, and
whether the candidate and baseline used the same load profile.

### 9. Budget gates

When a scenario has an authoritative budget file, the strict run must use it. FolderView strict
example:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -Configuration Release -CaseFilter folderView_perf_scroll_render_stress -PerfBudgetPath Specs\Testing\FolderViewPerfBudgets.json5 -RequirePerfBudgets -TimeoutMultiplier 8
```

Unknown machines must be visible: warn with the current `machineHash` and scaffold a budget entry.
Strict runs fail when the budget path is missing, no current-machine entry exists, or no hard entry
matches the current build flavor.

### 10. Instrumentation safety

Instrumentation must not distort the hot path it measures:

- no trace-file writes inside render/draw/present hot paths,
- no high-frequency per-item JSONL spam when per-pass aggregates answer the question,
- no stopwatch-only measurements when a metric row should exist,
- no hidden static cache that suppresses metrics after `ConfigureJsonlOutput(...)`,
- no production delay knobs; selftest-only latency hooks live under
  `RedSalamander/SelfTest/Common/SelfTestLatencyHooks.h/.cpp` and compile to no-ops outside
  `ENABLE_TESTS`.

---

## Apply To The Flake-Convergence Plan

`Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md` must reference this plan for any
work that changes performance-sensitive behavior. In particular:

- replacing `local_index_core_snapshot_reload`'s `warmElapsedMs < 1000u` assert must preserve an
  advisory timing metric and archive before/after evidence.
- converting `folderView_perf_overlay_invalidation_stress` to fixed-sample collection must preserve
  `folder.frame.total_us`, `folder.frame.present_us`, `metricQuality`, and analyzer-ready artifacts.
- scaling raw waits must not hide real latency regressions; add advisory metrics for cases where wait
  duration is a useful drift signal.
- the per-case timing dashboard must distinguish runner wall time from subsystem perf metrics.
- the nightly CPU-loaded timing lane must publish its environment matrix and load profile.

---

## Implementation Plan

### Task 0: Link The Contract

**Files:**
- Modify: `Specs/Plans/WIP/README.md`
- Modify: `Specs/Plans/WIP/Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md`
- Verify: `Specs/Testing/Testing_PerformanceValidation.md`

- [x] **Step 1: Add this plan to the WIP index**

Expected WIP index row:

```markdown
| `Operation_PerfMeasurementContract_2026-07-06.md` | Operational contract for scenario, metric keys, deterministic validation, archived evidence, analyzer quality gates, baseline/candidate comparison, and environment matrix on perf-sensitive work. | **Phase 0**: link the flake-convergence plan to this contract, then add representative Perf Measurement Records for overlay, local index snapshot reload, FileOps teardown/watchdog, and remote reachability probe. |
```

- [x] **Step 2: Link the flake-convergence plan**

Expected related-doc link in the flake-convergence plan:

```markdown
`Tools/Run-AllTests.ps1`, `Specs/Plans/WIP/Operation_PerfMeasurementContract_2026-07-06.md`.
```

- [x] **Step 3: Keep authority in the testing spec**

Confirmed: `Specs/Testing/Testing_PerformanceValidation.md` remains the normative spec; this WIP is
the execution plan and closeout checklist.

### Task 1: Install The Perf Measurement Record Template

**Files:**
- Modify: `Specs/Testing/Testing_PerformanceValidation.md`
- Modify: `Specs/Testing/Testing_SelfTests.md`
- Modify: `Specs/Plans/WIP/Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md`
- Test: `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`

- [ ] **Step 1: Write the failing documentation contract test**

Add a focused Pester test that fails if the authoritative spec stops exposing the required template
fields:

```powershell
It 'keeps the required Perf Measurement Record fields in the authoritative perf spec' {
    $repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    $spec = Get-Content -LiteralPath (Join-Path $repoRoot 'Specs\Testing\Testing_PerformanceValidation.md') -Raw
    foreach ($field in @(
        'Scenario:',
        'Metric keys:',
        'Deterministic validation:',
        'Build flavor:',
        'Baseline run:',
        'Candidate run:',
        'Analyzer command:',
        'Sample-quality result:',
        'Environment matrix:',
        'Archived evidence:',
        'Before/after delta:'
    )) {
        $spec | Should -Match ([regex]::Escape($field))
    }
}
```

- [ ] **Step 2: Run the failing test**

Run:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru"
```

Expected: FAIL if any required field is absent from the authoritative spec.

- [ ] **Step 3: Add the template to the authoritative spec**

Add the full template from this plan's `Contract` section to
`Specs/Testing/Testing_PerformanceValidation.md`, and keep this rule verbatim:

```markdown
A perf-sensitive change is not complete if the record has no scenario, no metric keys, no deterministic validation, or no archived evidence/blocked reason.
```

- [ ] **Step 4: Add representative records to the flake-convergence plan**

Under the timing/perf phase, add four concrete Perf Measurement Records:

```markdown
### Perf Measurement Record - FolderView overlay fixed-sample conversion

- Scenario: FolderView incremental-search overlay plus busy/cancel overlay animation under deterministic rendering load.
- Subsystem: FolderView Commands selftest.
- Change type: stabilization
- User-visible risk protected: overlay animation and repaint load must not turn CI load into false frame-sample failures while still catching render latency regressions.
- Metric keys: `folder.frame.total_us`, `folder.frame.present_us`, `metricQuality.folderFrameTotal.count`, `metricQuality.folderFramePresent.count`.
- Metric units and sample grain: microseconds per frame; one `metricQuality` summary per case run.
- Existing instrumentation reused or new instrumentation added: reuse existing FolderView frame metrics and `metricQuality`.
- Deterministic validation: `.\Tools\Run-AllTests.ps1 -Suite Commands -Configuration Release -CaseFilter folderView_perf_overlay_invalidation_stress -PerfBudgetPath Specs\Testing\FolderViewPerfBudgets.json5 -RequirePerfBudgets -TimeoutMultiplier 8`
- Build flavor: test-enabled Release for final evidence; Debug allowed only for diagnostic reproduction.
- Baseline run: [blocked] capture before conversion on the same machine.
- Candidate run: [blocked] capture after conversion on the same machine.
- Same-machine and same-suite comparison: [blocked] until both archives exist.
- Analyzer command: `.\Tools\Show-PerfRuns.ps1 -CompareRun <baseline>,<candidate> -FolderViewPreset -FailOnQuality`
- Sample-quality result: [blocked] until analyzer output exists.
- Environment matrix: [blocked] emitted by artifact; must include DPI, refresh rate, scale, local/RDP, WARP, adapter, OS, timeout multiplier.
- Archived evidence: [blocked] `Specs/TestRuns/<MachineHash>/Commands/<RunId>/perf-analysis.md`
- Before/after delta: [blocked] fill from analyzer table.
- Caveats or blockers: final evidence requires a test-enabled Release build.
- Authoritative spec updates required: update `Testing_PerformanceValidation.md` or `Testing_TestCoverage.md` if sample mode or metric keys change.

### Perf Measurement Record - local index snapshot reload timing

- Scenario: Local index core snapshot reload after cache eviction and small-tree search warmup.
- Subsystem: Search and index selftest.
- Change type: stabilization
- User-visible risk protected: a correct warm reload must not fail because Defender or runner load pushes an arbitrary elapsed-time ceiling over 1000 ms.
- Metric keys: `compare.selftest.local_index.snapshot_reload_us`.
- Metric units and sample grain: microseconds per snapshot reload; one row per measured reload.
  `value0` records candidate count and `value1` records snapshot file bytes.
- Existing instrumentation reused or new instrumentation added: new advisory `Debug::Perf::EmitDurationUs(...)` row emitted by `local_index_core_snapshot_reload` after the warm indexed query.
- Deterministic validation: final evidence uses
  `REDSALAMANDER_TEST_ROOT=C:\RSPerf` plus
  `.\Tools\Run-AllTests.ps1 -Suite Compare -Configuration Release -CaseFilter local_index_core_snapshot_reload -TimeoutMultiplier 8`.
  Diagnostic proof on this machine initially used
  `REDSALAMANDER_TEST_ROOT=C:\Users\eric\AppData\Local\Temp\RedSalamander-TestSandbox` because the
  default `Z:\src\RedSalamander\.build\TestSandbox` root is not NTFS; do not reuse that long profile
  path for final evidence.
- Build flavor: test-enabled Release for final evidence; Debug allowed only for diagnostic reproduction.
- Baseline run: [blocked] capture before replacing the elapsed ceiling.
- Candidate run: Debug diagnostic candidate captured in
  `Specs\TestRuns\4cb089111a23\CompareDirectories\2026-07-06_131232`.
- Same-machine and same-suite comparison: [blocked] baseline archive was not captured before the elapsed ceiling was replaced; candidate-only diagnostic evidence is directional.
- Analyzer command: `.\Tools\Show-PerfRuns.ps1 -Run .\Specs\TestRuns\4cb089111a23\CompareDirectories\2026-07-06_131232 -Metric compare.selftest.local_index.snapshot_reload_us -ShowBuildFlavor`
- Sample-quality result: one advisory sample, 6687 us, count=1; no p95/p99 claim because sample quality fails by design for a single run.
- Environment matrix: candidate artifact emitted under machine hash `4cb089111a23`; diagnostic root override used an NTFS `C:` sandbox.
- Archived evidence: `Specs\TestRuns\4cb089111a23\CompareDirectories\2026-07-06_131232\perf\perf_metrics.jsonl`
- Before/after delta: [blocked] fill from analyzer table or advisory metric summary.
- Caveats or blockers: final Release evidence and baseline/candidate delta remain blocked until a same-machine baseline is available; default `Z:` sandbox cannot exercise this NTFS-gated case.
- Authoritative spec updates required: document the advisory metric and remove any claim that a fixed wall-clock ceiling is a correctness gate.

### Perf Measurement Record - FileOps teardown/watchdog timing

- Scenario: File Operations teardown, cancel, and watchdog paths under deterministic slow-operation coverage.
- Subsystem: FileOperations selftest.
- Change type: stabilization
- User-visible risk protected: teardown and cancel must remain bounded without turning a loaded runner into a false watchdog failure.
- Metric keys: `FileOps.*cancel*`, `FileOps.*watchdog*`, `FileOps.*teardown*`, or the exact existing FileOps timing keys emitted by `Phase5_PreCalcCancelLatencyLocal`, `Phase5_CancelQueuedTask`, `Phase14_PopupHostLifetimeGuard`, `Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker`, or `Riptide_HostPerItemSchedulerShutdownWaitsForBlockedWorker`.
- Metric units and sample grain: microseconds or milliseconds per operation phase; one row per cancel/watchdog/teardown phase.
- Existing instrumentation reused or new instrumentation added: reuse current FileOps pre-calc/cancel metrics when they cover the touched path; add a named phase metric for uncovered teardown waits.
- Deterministic validation: `.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Release -CaseFilter Phase5_PreCalcCancelLatencyLocal,Phase5_CancelQueuedTask,Phase14_PopupHostLifetimeGuard,Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker,Riptide_HostPerItemSchedulerShutdownWaitsForBlockedWorker -TimeoutMultiplier 8`
- Build flavor: test-enabled Release for final evidence; Debug allowed only for diagnostic reproduction.
- Baseline run: [blocked] capture before wait-scaling or watchdog-policy change.
- Candidate run: [blocked] capture after wait-scaling or watchdog-policy change.
- Same-machine and same-suite comparison: [blocked] until both archives exist.
- Analyzer command: `.\Tools\Show-PerfRuns.ps1 -CompareRun <baseline>,<candidate> -Metric <exact FileOps metric key emitted by the touched case> -FailOnQuality`
- Sample-quality result: [blocked] report sample count and avoid p95 claims when the case is one-shot.
- Environment matrix: [blocked] include timeout multiplier and any injected slow-provider/load profile.
- Archived evidence: [blocked] `Specs/TestRuns/<MachineHash>/FileOps/<RunId>/perf-analysis.md`
- Before/after delta: [blocked] fill from analyzer table or one-shot phase comparison.
- Caveats or blockers: watchdog smoke limits are correctness boundaries; do not relax them unless deterministic evidence proves the old boundary was invalid.
- Authoritative spec updates required: update FileOps coverage notes if metric keys or watchdog expectations change.

### Perf Measurement Record - remote reachability probe timing

- Scenario: Compare Directories or remote-provider connection probe on credentialed but unreachable endpoints.
- Subsystem: Compare Directories or provider-specific remote selftest.
- Change type: stabilization
- User-visible risk protected: remote reachability probing must fail fast enough for users while avoiding CI-only timeout flakes from network or credential state.
- Metric keys: `compare.remote_probe.*`, provider-specific reachability timing keys, or a new metric that records probe start-to-result latency and outcome for `remote_file_s3`, `remote_s3_pagination`, `remote_file_ftp`, `remote_file_onedrive_personal`, `remote_file_onedrive_business`, or `remote_file_sharepoint`.
- Metric units and sample grain: milliseconds per probe attempt; one row per endpoint probe.
- Existing instrumentation reused or new instrumentation added: [blocked] confirm current metric coverage for remote probe latency; add one metric row if absent.
- Deterministic validation: `.\Tools\Run-AllTests.ps1 -Suite Compare -Configuration Release -CaseFilter remote_file_s3,remote_s3_pagination,remote_file_ftp,remote_file_onedrive_personal,remote_file_onedrive_business,remote_file_sharepoint -TimeoutMultiplier 8`
- Build flavor: test-enabled Release for final evidence; Debug allowed only for diagnostic reproduction.
- Baseline run: [blocked] capture before timeout/retry-policy change.
- Candidate run: [blocked] capture after timeout/retry-policy change.
- Same-machine and same-suite comparison: [blocked] until both archives exist.
- Analyzer command: `.\Tools\Show-PerfRuns.ps1 -CompareRun <baseline>,<candidate> -Metric <exact remote probe metric key emitted by the touched case> -FailOnQuality`
- Sample-quality result: [blocked] use deterministic synthetic endpoint samples; do not cite noisy live-network p95 as final evidence.
- Environment matrix: [blocked] include network fixture, credential mode, timeout multiplier, and local/RDP status.
- Archived evidence: [blocked] `Specs/TestRuns/<MachineHash>/CompareDirectories/<RunId>/perf-analysis.md`
- Before/after delta: [blocked] fill from analyzer table or deterministic probe summary.
- Caveats or blockers: live external endpoints are not acceptable as the only validation path.
- Authoritative spec updates required: document the deterministic remote probe fixture and metric family in the owning provider or Compare spec.
```

- [ ] **Step 5: Run the documentation contract test again**

Run:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru"
```

Expected: PASS for the Perf Measurement Record field guard.

### Task 2: Add Runner Perf Metadata To Aggregate Summaries

**Files:**
- Modify: `Tools/TestRunPlan.ps1`
- Modify: `Tools/Run-AllTests.ps1`
- Test: `Tools/Tests/RunAllTestsPlan.Tests.ps1`
- Document: `Specs/Testing/Testing_SelfTests.md`

- [ ] **Step 1: Write the failing summary-schema test**

Add a Pester test that calls `New-RSTestRunSummary` with perf metadata and verifies the serialized
shape:

```powershell
It 'records perf measurement metadata in aggregate run summaries' {
    $result = [pscustomobject]@{
        Name = 'Commands'
        Type = 'SelfTest'
        Passed = $true
        ExitCode = 0
        ResultsPath = 'Z:\tmp\results.json'
        TracePath = 'Z:\tmp\trace.txt'
    }

    $summary = New-RSTestRunSummary `
        -Suite 'Commands' `
        -Configuration 'Release' `
        -Results @($result) `
        -StartTime ([datetime]'2026-07-06T10:00:00Z') `
        -EndTime ([datetime]'2026-07-06T10:01:00Z') `
        -TestRoot 'Z:\src\RedSalamander\.build\TestSandbox' `
        -RunId '20260706T100000Z_1234_abcd' `
        -PerfMetadata ([pscustomobject]@{
            BuildFlavor = 'Release'
            MachineHash = 'abc123def456'
            TimeoutMultiplier = 8.0
            ArchivePath = 'Specs\TestRuns\abc123def456\Commands\2026-07-06_100000'
            AnalyzerCommand = '.\Tools\Show-PerfRuns.ps1 -Run <run> -FolderViewPreset -FailOnQuality'
        })

    $summary.performance.build_flavor | Should -Be 'Release'
    $summary.performance.machine_hash | Should -Be 'abc123def456'
    $summary.performance.timeout_multiplier | Should -Be 8.0
    $summary.performance.archive_path | Should -Match 'Specs\\TestRuns'
    $summary.performance.analyzer_command | Should -Match 'Show-PerfRuns.ps1'
}
```

- [ ] **Step 2: Run the failing test**

Run:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru"
```

Expected: FAIL because `New-RSTestRunSummary` does not yet accept or serialize `PerfMetadata`.

- [ ] **Step 3: Implement `PerfMetadata` summary support**

Add an optional `-PerfMetadata` parameter to `New-RSTestRunSummary` in `Tools/TestRunPlan.ps1` and
serialize this object when supplied:

```powershell
performance = [pscustomobject]@{
    build_flavor       = $PerfMetadata.BuildFlavor
    machine_hash       = $PerfMetadata.MachineHash
    timeout_multiplier = $PerfMetadata.TimeoutMultiplier
    archive_path       = $PerfMetadata.ArchivePath
    analyzer_command   = $PerfMetadata.AnalyzerCommand
}
```

- [ ] **Step 4: Populate metadata in the runner**

In `Tools/Run-AllTests.ps1`, construct `PerfMetadata` from the actual configuration, timeout
multiplier, run context, and archive discovery before writing `run-all-tests-results.json`.

- [ ] **Step 5: Re-run the focused Pester suite**

Run:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru"
```

Expected: PASS with the new summary metadata test.

### Task 3: Preserve And Point To Archived Perf Evidence

**Files:**
- Modify: `Tools/TestRunPlan.ps1`
- Modify: `Tools/Run-AllTests.ps1`
- Modify: `Specs/TestRuns/README.md`
- Test: `Tools/Tests/RunAllTestsPlan.Tests.ps1`

- [ ] **Step 1: Write the failing archive-path test**

Add a test for a pure helper that resolves the intended archive destination:

```powershell
It 'formats perf archive destinations by machine, area, and run id' {
    $path = Get-RSPerfArchiveDestination `
        -RepoRoot 'Z:\src\RedSalamander' `
        -MachineHash 'abc123def456' `
        -Area 'Commands' `
        -RunId '2026-07-06_100000'

    $path | Should -Be 'Z:\src\RedSalamander\Specs\TestRuns\abc123def456\Commands\2026-07-06_100000'
}
```

- [ ] **Step 2: Run the failing test**

Run:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru"
```

Expected: FAIL because `Get-RSPerfArchiveDestination` does not exist.

- [ ] **Step 3: Implement the pure archive destination helper**

Add `Get-RSPerfArchiveDestination` to `Tools/TestRunPlan.ps1` using `Join-Path`, not string
concatenation, and keep the return value as a fully resolved path.

- [ ] **Step 4: Preserve partial artifacts on timeout/crash**

Update `Run-AllTests.ps1` so `results.json`, `trace.txt`, and `perf_metrics.jsonl` are copied or
reported when a harness exits non-zero, times out, or returns partial output. Do not treat partial
perf artifacts as pass evidence; only preserve them for diagnosis.

- [ ] **Step 5: Document the archive rule**

In `Specs/TestRuns/README.md`, require `perf-analysis.md` beside archived perf artifacts whenever
the run supports a before/after claim.

- [ ] **Step 6: Re-run focused tests**

Run:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru"
```

Expected: PASS with archive destination and partial-artifact coverage.

### Task 4: Gate Analyzer Quality

**Files:**
- Modify: `Tools/Show-PerfRuns.ps1`
- Test: `Tools/Tests/ShowPerfRuns.Tests.ps1`
- Document: `Specs/Testing/Testing_PerformanceValidation.md`

- [ ] **Step 1: Add the failing low-sample p95/p99 test**

Add a test that invokes `Show-PerfRuns.ps1 -FailOnQuality` against a fixture with too few samples for
a requested percentile metric.

```powershell
$result = & $pwsh -NoProfile -ExecutionPolicy Bypass -File $showPerfRunsScript -Run $runPath -Metric 'folder.frame.total_us' -FailOnQuality 2>&1
$LASTEXITCODE | Should -Not -Be 0
($result -join "`n") | Should -Match 'P95Quality'
```

- [ ] **Step 2: Run the failing or tightened analyzer test**

Run:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\ShowPerfRuns.Tests.ps1 -PassThru"
```

Expected: FAIL before the quality gate rejects low-sample percentile claims.

- [ ] **Step 3: Implement or tighten the quality gate**

Ensure `-FailOnQuality` exits non-zero when a requested p95/p99 metric lacks enough samples, while
still allowing non-quality-gated informational rows to print.

- [ ] **Step 4: Re-run analyzer tests**

Run:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\ShowPerfRuns.Tests.ps1 -PassThru"
```

Expected: PASS.

### Task 5: Emit Environment Matrix Metadata

**Files:**
- Modify: `Tools/TestRunPlan.ps1`
- Modify: `Tools/Run-AllTests.ps1`
- Modify: native selftest artifact writers only if PowerShell cannot observe a required field
- Test: `Tools/Tests/RunAllTestsPlan.Tests.ps1`
- Document: `Specs/Testing/Testing_PerformanceValidation.md`

- [ ] **Step 1: Write the failing environment-matrix test**

```powershell
It 'records the required perf environment matrix fields' {
    $matrix = New-RSPerfEnvironmentMatrix -BuildFlavor 'Release' -TimeoutMultiplier 8.0

    foreach ($name in @(
        'buildFlavor',
        'osVersion',
        'timeoutMultiplier',
        'activeDpi',
        'displayRefreshHz',
        'displayScalePercent',
        'isRemoteSession',
        'warpAvailable',
        'warpRunExecuted',
        'adapterName',
        'driverVersionAvailable'
    )) {
        $matrix.PSObject.Properties.Name | Should -Contain $name
    }
}
```

- [ ] **Step 2: Run the failing test**

Run:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru"
```

Expected: FAIL because `New-RSPerfEnvironmentMatrix` does not exist.

- [ ] **Step 3: Implement the matrix helper**

Add `New-RSPerfEnvironmentMatrix` in `Tools/TestRunPlan.ps1`. Use PowerShell-observable values for
OS version, remote-session status, timeout multiplier, and environment flags. Use explicit
`$null`/`'unknown'` values for fields that require native DxGI/DPI capture until the native harness
emits them.

- [ ] **Step 4: Attach the matrix to perf metadata**

Add the matrix to `run-all-tests-results.json` under `performance.environment_matrix`, and document
which values are runner-observed versus native-observed.

- [ ] **Step 5: Re-run focused tests**

Run:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru"
```

Expected: PASS with environment-matrix metadata.

### Task 6: Enforce Closeout Before Moving Plans To Done

**Files:**
- Modify: `Specs/Testing/Testing_PerformanceValidation.md`
- Modify: `Specs/Plans/WIP/README.md`
- Test: `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`

- [ ] **Step 1: Add a closeout checklist guard**

Document this exact closeout rule in the WIP index and authoritative perf spec:

```markdown
Perf-sensitive WIP plans cannot move to `Specs/Plans/Done/` until every perf-sensitive task has a Perf Measurement Record, archived evidence or a blocked owner/task, and any durable metric/scenario rule has been merged into the authoritative subsystem spec.
```

- [ ] **Step 2: Add a source-contract test for automated lint**

Add a Pester guard that checks WIP/Done plans mentioning `perf`, `p95`, `p99`, `latency`,
`throughput`, `frame`, or `timing` also mention `Perf Measurement Record` or link to
`Testing_PerformanceValidation.md`.

- [ ] **Step 3: Run the guard**

Run:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru"
```

Expected: PASS once every active perf-sensitive plan has an explicit record or authoritative link.

---

## Audit Checklist

For each WIP plan or PR that touches perf-sensitive code:

1. Does it name the scenario before implementation?
2. Does it list metric keys and units?
3. Does it reuse or add instrumentation?
4. Does it include deterministic validation commands?
5. Does it state build flavor and use test-enabled Release where required?
6. Does it archive `results.json`, `trace.txt`, and `perf_metrics.jsonl` under `Specs/TestRuns/`?
7. Does it compare baseline and candidate on the same machine or label the result directional?
8. Does it run `Tools/Show-PerfRuns.ps1 -FailOnQuality` for percentile claims?
9. Does it include an environment matrix?
10. Does it update the authoritative subsystem spec before the WIP plan moves to Done?

---

## Exit Criteria

- The flake-convergence plan links to this plan for timing/perf-sensitive work.
- At least four representative Perf Measurement Records exist: FolderView overlay, local index
  snapshot reload, FileOps teardown/watchdog, and remote reachability probe.
- `Run-AllTests.ps1` artifacts and archived `Specs/TestRuns/` folders carry the required metadata.
- Percentile claims are analyzer-gated with `-FailOnQuality`.
- The WIP README points to this plan as the owner for perf-measurement contract rollout.
- Durable requirements discovered here are merged into `Testing_PerformanceValidation.md`,
  `Testing_SelfTests.md`, and affected subsystem specs.
- This plan moves to `Specs/Plans/Done/` only after the above are complete.

---

## Quick Wins

1. Add the Perf Measurement Record template to the flake-convergence plan's Phase 3 section.
2. Require `Specs/TestRuns/<MachineHash>/<Area>/<RunId>/perf-analysis.md` for overlay and snapshot
   reload fixes.
3. Make `Show-PerfRuns.ps1 -FailOnQuality` mandatory for any p95/p99 statement in closeout.
4. Add build flavor and `machineHash` to the per-case timing/flake dashboard rows.
5. Define the CPU-loaded runner profile before using "timing robustness" as exit evidence.
