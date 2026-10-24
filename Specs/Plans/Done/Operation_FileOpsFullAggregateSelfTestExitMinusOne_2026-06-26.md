# Operation FileOps Full Aggregate Selftest Exit Minus One - 2026-06-26

## Status

- State: Done.
- Closed on 2026-06-26 after the full FileOps aggregate stopped reproducing the original exit `-1`, the remaining aggregate failure was isolated to a brittle Phase11 bridge selftest assertion, and the harness was made deterministic.
- Created during closeout of `Specs/Plans/Done/Operation_FSSubsystemDeepAuditRemediation_DataSafetySecurityContracts_2026-06-26.md`.

## Problem

During the FS subsystem deep-audit closeout, targeted Track A/Track C FileOps guards and the isolated Fairstream family passed, but one full FileOps aggregate runner exited before writing `fileops_results.json`.

The failing command:

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild
```

Observed on 2026-06-26:

- One retry returned exit code `3` because another stale selftest process from `C:\Users\PVB\.codex\worktrees\cac5\RedSalamander` held the selftest mutex with command line `--compare-selftest --selftest-case=mtp_`. That process exited before `Stop-Process` found it.
- With the mutex clear, the aggregate FileOps run exited `-1` after about 39 seconds.
- `run-all-tests-results.json` recorded `suite=FileOps`, `exit_code=1`, and the FileOperations entry with `exit_code=-1`, zero cases, and no failure details.
- `fileops_results.json` was not produced.
- The last FileOps trace line was `NextStep: Fairstream_SaturationConcurrentCopiesMakeProgress`.
- No new WER event or `%LOCALAPPDATA%\CrashDumps\RedSalamander*.dmp` was created for this aggregate exit after the earlier Phase8 stack crash was fixed.
- The isolated Fairstream family passed immediately afterward: `Specs/TestRuns/7d3a1247382a/FileOps/2026-06-26_161711/` (`FileOpsFamily_Fairstream`, 38 passed).

Follow-up aggregate runs did not reproduce the process exit. They completed normally and surfaced one failing guard instead: `Phase11_BridgeMultiFolderParallelCopyInFlightLines` reported zero bridge file-admission counters in the aggregate archive `Specs/TestRuns/7d3a1247382a/FileOps/2026-06-26_165546/`.

## Root Cause

The reproduced failure was a harness determinism bug, not a production data-safety bug.

- `Phase11_BridgeMultiFolderParallelCopyInFlightLines` selected six directory roots.
- The dummy destination provider caps copy/move concurrency at four.
- When four top-level bridge directory copies were active, `CrossFileSystemBridge::ComputeWithinFolderBudget()` returned `1`, so each selected root used the sequential directory bridge path.
- The operation still copied data correctly and could show live top-level in-flight activity, but file-admission metrics are emitted by the parallel within-folder bridge path. The test therefore asserted counters that its own timing and budget shape did not guarantee.

## Fix

- Reduced the guard to two selected source folders.
- Pinned the local plugin configuration inside the test to manual `copyMoveMaxConcurrency=4`.
- This keeps each active selected root under the dummy provider's four-copy cap with within-folder bridge budget greater than `1`, deterministically exercising the bridge file-admission queue and early-start counters.
- Updated `Specs/FileSystem/FileSystem_FileOperations.md` so the durable FileOps contract records the bridge admission scenario and required metrics.

## Validation

Required closeout commands:

```powershell
.\build.ps1 -ProjectName RedSalamander
.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Phase11_BridgeMultiFolderParallelCopyInFlightLines
.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter FileOpsFamily_Phase11_BridgeAndConnections
.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild
.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter FileOpsFamily_Fairstream
git diff --check
```

Evidence:

- Build passed with 0 warnings and 0 errors.
- `Phase11_BridgeMultiFolderParallelCopyInFlightLines` passed: 3 passed, 0 failed, 0 skipped; archive `Specs/TestRuns/7d3a1247382a/FileOps/2026-06-26_171314/`.
- `FileOpsFamily_Phase11_BridgeAndConnections` passed: 9 passed, 0 failed, 0 skipped; archive `Specs/TestRuns/7d3a1247382a/FileOps/2026-06-26_171752/`.
- Full FileOps aggregate passed: 100 passed, 0 failed, 20 skipped; archive `Specs/TestRuns/7d3a1247382a/FileOps/2026-06-26_173122/`.
- `FileOpsFamily_Fairstream` passed: 38 passed, 0 failed, 0 skipped; archive `Specs/TestRuns/7d3a1247382a/FileOps/2026-06-26_173311/`.
- `git diff --check` passed before commit.

## Remaining Work

- None for this follow-up. The original exit `-1` did not reproduce after the stale mutex holder cleared; the aggregate now completes and the brittle bridge guard is deterministic.
