# C0 Curl partial-tree failure — passing Debug runtime

- Source: `58f22d0e` plus the recorded C0 changes, explicit selected-item mutation
  proof, expanded controls, and corrected first-listing refusal placement.
- Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_ -FailFast -RunId c0-curl-partial-tree-qualified-debug-20260905`.
- Build receipt: `5413ea7b5d3973dee511f0d43a5f3d9e21f6def13c32bc96b884b4068a143cb6`;
  build 2:57, zero compiler errors, **two MSB3061 cleanup warnings**. This is not
  warning-free build qualification; `build.txt` preserves the complete log.
- Harness: 11 passed, zero failed/skipped. Native guards 29/29, lost replies 33/33,
  partial-tree/proof/collision checks 37/37.

The original serial partial-tree failures now report unknown instead of inventing
no-commit proof. Both discovery-first controls remain untouched and retryable.
Two selected items retain independent proof at concurrency 1/4. All four native
Move/Rename collision controls preserve source/occupant bytes with zero mutation
commands and explicit safe-retry proof. Reporter controls reject bare-error proof.

The real host partial-tree witness passed in 781,000 us: two mutation commands,
one deleted leaf, preserved later child/sibling, immutable Indeterminate result,
localized unknown summary and zero conflict prompts. This verifies the existing
summary projection, not a manual screen-reader or full UX review.

Same-machine partial-tree timings, single serial/parallel then bulk serial/parallel:
5,858 / 5,566 / 5,216 / 5,006 us, versus pre-fix 8,297 / 4,277 / 5,080 / 4,889 us.
The purpose is correct completion proof; mixed millisecond changes are not a claimed
speedup. Both profiles retain the exact mutation-count invariant. The independent
item, collision and host controls were added after the original baseline.

Raw results, trace, aggregate and complete perf stream are copied intact from this
run's `artifacts/selftest/last_run` beneath `C:\RedSalamander.Perf\runs`.
Disk audit: 40 retained-run issues, nothing deleted.

The cleanup warnings name shared `Plugins/libcrypto-3-x64.dll` and `sqlite3.dll`.
Several per-project FileListAbsolute records list the same output DLL: Curl and
Google Drive list libcrypto; Local and ViewerSqlite list sqlite3; many plugins list
yyjson (the earlier build warning). MSBuild CoreClean deletes those per-project
prior outputs. This ownership overlap is a concrete closeout investigation target,
not a proven complete diagnosis of every transient access denial. No process was
killed. Native walker coverage, C1–C6 and final Fresh Full remain open.
