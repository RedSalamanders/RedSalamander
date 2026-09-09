# File Operations Remaining Stress Validation

> **FILE OPERATIONS CONTRACT (2026-08-21).** New evidence uses the finalized metric dictionary in `Specs/Testing/Testing_PerformanceValidation.md`: endpoint topology, selected strategy, discovery bounds, verification bytes, digest algorithm/CPU, verification outcome, and link kind/policy/action. "Same filesystem" alone is invalid.


> **NON-NORMATIVE ACTIVE PLAN.** Current behavior is defined by
> `Specs/FileSystem/FileSystem_FileOperations.md`; this file owns only the
> remaining validation/evidence work recovered from its retired implementation
> appendices.

## Status

- **State:** ACTIVE
- **Priority:** P2 validation debt
- **Planned at:** `e14ba96e30d7f767b4020f271f07a0936fe56b7d`
- **Ownership boundary:** missing deterministic stress and visual evidence for
  already-implemented File Operations behavior; no new feature semantics
- **Drift check:**
  `git diff e14ba96e30d7f767b4020f271f07a0936fe56b7d..HEAD -- RedSalamander Plugins/FileSystem Common/PlugInterfaces/FileSystem.h Specs/FileSystem/FileSystem_FileOperations.md Tests Tools`

## Existing baseline

The File Operations selftest already covers watcher churn, concurrency knobs,
recursive single- and multi-root copy, the recursive shape matrix, speed-limit
changes, in-flight progress streams, bridge admission, basic parallel delete,
popup resize/pause smoke, and destructive correctness. Current cases are
enumerated by `Tools/Get-TestInventory.ps1`; this plan does not duplicate their
full names as a frozen inventory.

## Remaining outcomes

1. Add deterministic popup-model/render assertions for parallel in-flight rows,
   graph scaling, paused/waiting/calculating overlays, rainbow sample stability,
   and narrow-width middle ellipsis.
2. Exercise directory enumeration with many long names and deliberately small
   soft/hard caps; prove the expected hard-cap failure and post-run buffer trim.
3. Record same-machine many-small and several-large file evidence for copy/move
   concurrency `1`, `4`, and `16`, including mid-flight speed-limit changes and
   repeated pause/resume convergence.
4. Exercise parallel delete over deep trees with Recycle Bin on/off and the
   supported delete/recycle concurrency knobs. Environment-dependent Recycle
   Bin behavior must skip with an explicit reason rather than claim coverage.
5. Exercise Queue/Parallel switching with multiple tasks in pre-calculation and
   multiple tasks in transfer; prove only the intended task advances in Queue
   mode and all eligible tasks resume in Parallel mode.

## Files in scope

- `RedSalamander/FolderWindow.FileOperations*.cpp`
- `Plugins/FileSystem/FileSystem.FileOps.cpp`
- `Plugins/FileSystem/FileSystem.DirectoryOps.cpp`
- File Operations selftests and deterministic popup model tests
- `Specs/TestRuns/<commit>/FileOps/` for archived before/after or validation
  evidence

## Verification

```powershell
.\build.ps1 -ProjectName FileSystem
.\.build\x64\Debug\RedSalamander.exe --fileops-selftest
.\Tools\Run-AllTests.ps1 -Suite Full
```

Every perf-sensitive row also follows
`Specs/Testing/Testing_PerformanceValidation.md`, including scenario identity,
instrumentation, reproducible capture, correctness gates, and archived evidence.

## Done criteria

- All five outcomes have deterministic coverage or an explicit, reviewed
  environment skip with a reproducible manual procedure.
- Focused and Full gates are green.
- Durable test/performance obligations are merged into the current File
  Operations and testing specs.
- This plan is moved to `Specs/Plans/Done/` only after those gates pass.

## STOP conditions

Stop and request a new product/architecture decision if a test would require a
different conflict policy, queue order, bandwidth meaning, public plugin ABI,
or destructive-operation rule. A validation plan cannot redefine behavior.
