# File Operations checkpoint: fresh x64 Release PR qualification

The test-enabled x64 Release solution builds with zero compiler warnings/errors.
The subsequent **fresh PR gate passes: 13 entries executed and promoted, 824 reported
cases passed, zero failed/skipped, and zero sandbox-audit issues**. No entry was reused,
provisional, invalidated or incomplete. No quarantine retry was needed. The source
snapshot is unchanged throughout the run.

The runner's 824 reported cases comprise 812 Pester cases and twelve native/script
entry outcomes. Detailed assertions inside those native suites remain in the full
logs; the aggregate does not count each of them as a separate case.

This is the requested **PR gate**, not Full closeout or acceptance of the outstanding
DxUi menu/grid candidates. All thirteen entries declare `RequiresInteractiveDesktop=false`;
the driver checks this before launch. No foreground inspection, mixed-DPI journey,
assistive-technology acceptance or paired resource comparison is claimed here.

## Reproduction and identity

| Item | Recorded value |
| --- | --- |
| Source | `ccebd5af68eb45f2e9deed64a7063780e71eb1ce`, clean detached `C:/RsI26Final` |
| Consumer DxUi pin | `78b3de389a189c7f86f611787e0489fb6d474218` |
| vcpkg tool checkout | `1391150599c0b9dc6f2a3c79c50eff6c94b077d1`, `C:/RsI26Vcpkg` |
| Build | `build.ps1 -Configuration Release -Platform x64 -MaxCpuCount 4`, tests enabled |
| Test command | `Tools/Run-AllTests.ps1 -Suite PR -Configuration Release -Platform x64 -SkipBuild -TestRoot E:/RedSalamander.Perf` |
| Run | `20260923T082913Z-68892-de68d74a1fba4795a8967d2962bd0365` |
| Build receipt | `09848e19caf0b612e0663dae09d6079c17255b4862d4bd955e7ebed7b86ed1f1` |
| Snapshot | `c17668cccb378dd7bba5e348658ad557a3078098de56b82c6e2359e23da22cd5` |
| Plan | `96502d33b8c2c3ae363e846a81b2a2e3c4e78d31737af68174ab7a025b6de53f` |
| Decision | `548ceef297ff33ae971720f35d9dc407cd80a6c0ae94a1142c1ae94592373eec` |
| Coverage plan | `69e28753fc743bb976101c5232cd2af218a81fbe15adea12d50ca86db5367354` |
| Build / test duration | 471,268 ms / 452,476 ms |
| Exit codes | Build 0, PR 0 |

The [exact driver](raw/qualify-release.ps1.receipt.txt) records the process-local CI
vcpkg imports/include/library paths. No global tool integration changed. The short
checkout preserves the independently launched primary Release application.

Dependency setup first rejected a Visual Studio vcpkg executable lacking the pinned
Git checkout, then encountered an unavailable historic 7zip tree in the shallow
pinned checkout. Fetching the tool checkout's full history resolved setup; all 35
dependencies subsequently installed. One upstream LibRaw CMake unused-variable warning
is retained in external dependency logs. The zero-warning claim above applies to the
product solution build, not every dependency setup attempt. The restore's announcement
of a newer remote DxUi main did not change the consumer pin.

## Evidence and limits

- [Aggregate v2](raw/run-all-tests-results.json) retains the requested gate, verdict
  axes, entry/origin identities, promotion digests and zero reused/avoided duration.
  `coverage_scope=complete` and `repository_verdict=PASSED` apply to **PR**.
- [Legacy case sidecar](raw/run-all-tests-results-v1.json) and the complete
  [ordered case history](raw/run-all-tests-case-history.jsonl) retain per-case status,
  classifications and audit details. The [dashboard](raw/run-all-tests-dashboard.md)
  provides a readable entry summary.
- [Immutable plan](raw/validation-plan.json), [decision](raw/validation-plan-decision.json),
  [build receipt](raw/build-receipt.json), initial/final workspace snapshots, and this
  run's [entry promotions, evidence and result artifacts](raw/validation/) preserve
  the executable and provenance chain. This is a fresh run, not composite evidence.
- Complete [product build](raw/build.log.receipt.txt) and [PR runner](raw/pr.log.receipt.txt)
  outputs are included. Raw bytes are preserved; [SHA256SUMS.txt](SHA256SUMS.txt) binds
  all 75 raw files. The approximately 2.46 MiB raw payload stays below the archive limits.
- Full dependency build/configure logs, the explicitly partial dependency console
  transcript, and the complete original run copies remain at
  `C:/RedSalamander.Perf/evidence/i26-ui/release-20260923/`. They are not a reusable-cache
  archive or a new performance baseline.

Remaining File Operations gates are tracked in the
[active plan](../../../../Plans/WIP/UI_FileOperationsReview_2026-09-16.md): the warned
foreground inspection rerun after desktop availability is confirmed, fresh Full on
the clean root, native DPI/accessibility acceptance, resource decisions/qualification,
explicit combined-pin adoption, and final integration/gallery/spec closeout. Existing
failures and unrelated owners' deferred work remain open.
