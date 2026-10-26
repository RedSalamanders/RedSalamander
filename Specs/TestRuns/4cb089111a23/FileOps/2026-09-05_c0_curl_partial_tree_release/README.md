# C0 Curl partial-tree failure — test-enabled Release

- Source: `58f22d0e` plus C0 working changes; snapshot
  `37858cb4eb796e8d3dd2c8ded2ee0276c1abb32913f61e649aad77fdff57a6ac`, identical to
  the passing C0 Debug and ordinary FTP Debug regression receipts.
- Command: with process `RSBuildEnableTests=true`,
  `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Release -CaseFilter C0_ -FailFast -RunId c0-curl-partial-tree-release-20260905`;
  the previous environment value was restored after the run.
- Receipt: `49fda44b8ae9be96f7abf0d3d306bee44a742c317d63fda7cae417fac492bda5`;
  x64 Release, tests enabled, zero warnings/errors, build 3:52.
- Runtime: 11 passed, zero failed/skipped. Native root guards 29/29, lost replies
  33/33, partial-tree/shared-proof/independent-item/collision checks 37/37.

Serial partial-tree deletion records two mutation commands and unknown completion;
discovery-first controls record zero commands and retain no-commit. Independent
selected-item controls remain truthful at concurrency 1/4. Native Move/Rename
collisions preserve both source and occupant with zero mutation requests.

Partial-tree timings (single serial/parallel, bulk serial/parallel):
3,453 / 3,439 / 3,940 / 3,802 us. The host partial-tree witness completed in
63,000 us with two commands, one deleted leaf, preserved later child/sibling,
immutable Indeterminate result, localized unknown summary and zero conflict
prompts. Release versus Debug timings are not an apples-to-apples speedup claim.

Raw results, trace, aggregate and complete perf stream were copied intact from
this run's `artifacts/selftest/last_run` under `C:\RedSalamander.Perf\runs`.
Disk audit reported 42 retained-run issues; nothing was deleted. This focused
test-enabled Release pass is not the non-test Release SQLite staging transition,
shared-output cleanup-race qualification, native-walker retirement, or Fresh Full.
