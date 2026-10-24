# Microsoft Drive Merge Destination-Metadata Reuse Optimization

**Status:** DONE - implementation, correctness, performance evidence, production guard, and durable specs complete
**Date:** 2026-07-21
**Completed:** 2026-07-21
**Parent:** `Specs/Plans/WIP/Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md` (`FG-P2-6`)
**Protected scenario:** same-drive Microsoft Drive recursive MOVE of a folder containing many no-conflict files into an existing destination folder

## Objective

Remove the second, behaviorally redundant destination-child metadata GET performed by `MoveOrRenameItem` after `MergeMoveFolderIntoExisting` has already completed the current child's conflict probe. Preserve source/path freshness, Retry re-probing, overwrite rollback, partial-result behavior, concurrent-change safety, and the existing sequential mutation order.

## Last Finding and Current Spec/Code State

FG-P2-6 is complete. `MergeMoveFolderIntoExisting` now passes its current child's explicit `Present`/`Missing` destination result into `MoveOrRenameItem`; direct callers still perform their normal lookup, and Retry replaces the hint from a fresh probe. The production move path emits aggregate metadata GET count/time and child count, while the test-enabled same-binary legacy scope preserves a deterministic baseline.

For a one-page, 16-file, no-conflict merge, the legacy call graph performs 83 measured item-metadata GETs / 84 total GETs / 16 PATCHes. The candidate repeatedly proves 67 / 68 / 16, removing exactly one serial GET per child. Five same-machine Release samples show median fixed-latency time of 1,323,954 us versus 1,059,139 us, a 264,815 us / 20.00% reduction.

The initial audit suggested also reusing source-listing and parent metadata. The final safety review narrows that direction: the current by-path source and parent lookups revalidate that the named source and destination parents still identify the user's requested paths after enumeration. Blindly moving a listed source ID, or continuing to use a parent ID after that folder moved or was replaced, can mutate the wrong current namespace. Those lookups remain in this slice. Microsoft Graph supports `If-Match` on item PATCH, but adopting conditional source-ID mutation requires a separate concurrency design and proof; it is not necessary to close the exact redundant-destination-GET row.

The merge remains intentionally serial. Graph `$batch`/proactive pacing is deferred: after duplicate removal, the remaining per-child destination probe is the just-in-time conflict/freshness boundary, and the safe candidate already yields a material 20.00% median improvement without changing mutation count or order. Batching would broaden the metadata snapshot window and complicate prompt ordering, per-subrequest throttling, cancellation, overwrite backup/rollback, and partial completion. Reopen only with measured live throttling or large-merge latency evidence plus a separate deterministic mixed-result contract.

The working tree also contains concurrent user-owned edits to `Specs/Plans/WIP/UI_ThemeCycleOverlayPlan_2026-07-21.md` and `Specs/Plans/WIP/README.md`; this plan will not touch or stage either file.

## Required Behavior Contract

- A merge child receives exactly one just-in-time destination metadata probe before its first conflict decision.
- `MoveOrRenameItem` may consume a merge-local destination lookup hint only when it is explicitly `Present` or `Missing`; direct/top-level callers retain their normal lookup.
- Retry always performs a new destination probe and replaces the hint before another mutation attempt.
- A destination created after a `Missing` probe must make the server mutation fail closed; it must not be silently overwritten or deleted.
- Existing-destination overwrite still uses the probed item ID for reversible backup and retains the cleanup-debt contract.
- Source item, source parent, and destination parent by-path freshness checks remain unchanged in this slice.
- Directory merge ordering, recursive descent, cancellation, prompt paths, source-folder retention, and partial-result rules remain unchanged.
- Production perf capture emits aggregate merge metadata counts/timing, not one JSONL row per Graph request.

## Measurement and Deterministic Proof

The Microsoft Drive fake Graph must be available in test-enabled Release builds while remaining absent from a normal tests-disabled Release DLL. The harness will add fixed deterministic GET latency and a test-only legacy-refetch scope so baseline and candidate run in the same binary against identical graphs.

Metric family:

- `FileOps.MicrosoftDrive.Merge.MetadataGetCount`
- `FileOps.MicrosoftDrive.Merge.MetadataGetUs`
- `FileOps.MicrosoftDrive.Merge.ChildCount`
- `FileOps.MicrosoftDrive.Merge.Baseline`
- `FileOps.MicrosoftDrive.Merge.Candidate`
- `FileOps.MicrosoftDrive.Merge.Improvement`

The 16-child scenario must assert exact baseline/candidate item-metadata GETs, total Graph GETs, PATCHes, result status, destination identities, source-child removal, retained source-folder cleanup debt, and zero unexpected prompts. Focused cases must also lock existing-destination overwrite, Retry-to-missing, nested folder merge, and a destination created after the first missing probe.

## Progress Log

- [x] Verified the open FG-P2-6 parent row against current code.
- [x] Counted the current no-conflict request shape from `MergeMoveFolderIntoExisting` and `MoveOrRenameItem`: one destination probe plus four metadata GETs in the nested move, per child.
- [x] Corrected the safe scope: reuse only the just-fetched destination lookup; retain source and parent path-identity refreshes.
- [x] Defined the metric family, deterministic same-binary baseline/candidate, focused concurrency contracts, dirty-worktree boundary, and Done gate before code changes.
- [x] Added aggregate logical metadata-GET count/time and merge-child instrumentation, production `FileOps.MicrosoftDrive.Merge.*` emission, fixed-latency fake-Graph support, exact PATCH collision behavior, and a same-binary legacy/candidate scenario.
- [x] Moved Microsoft Drive's test-only Graph transport/selftests from `_DEBUG` to the repository-standard `ENABLE_TESTS` gate so test-enabled Release can execute the provider contract while normal Release remains clean.
- [x] The first test-enabled Release compile found the shared `Common::DebugSelfTest::Check` declaration still gated by `_DEBUG`; aligned that test-only helper with the repository's `ENABLE_TESTS` contract. Initial failed log: `.build/logs/msbuild-20260721_123002_212.log` (103 cascade errors, all rooted at the missing helper type; 0 warnings).
- [x] The second test-enabled Release compile reached link and found that Microsoft Drive had no Common project reference, so the new canonical `Debug::Perf` calls could not resolve `PerfJsonl` exports. Added the same centralized Common project reference used by S3/Curl. Failed log: `.build/logs/msbuild-20260721_123116_068.log` (3 link errors, 0 warnings).
- [x] Wired explicit merge-local `Present`/`Missing` destination hints through missing, overwrite, and Retry paths, but deliberately left `MoveOrRenameItem` ignoring the hint at this RED checkpoint.
- [x] Test-enabled Release provider and `PluginContractTests` builds passed with 0 warnings/errors after the two harness-linkage corrections (`msbuild-20260721_123217_641.log`, `msbuild-20260721_123305_294.log`).
- [x] Captured the RED same-binary run under `.build/TestSandbox/runs/20260721T103500Z-msdrive-merge-red/`. Microsoft Drive reported 119 passed / 2 failed: the expected candidate exact-count and improvement assertions. Legacy and not-yet-optimized candidate both measured 83 logical metadata GETs / 84 total Graph GETs / 16 PATCHes; fixed-latency metadata time was 1,315,745 / 1,305,202 us. All transfer-result assertions passed. The two S3 export failures are unrelated artifact configuration: S3 had intentionally been overwritten by the preceding task's production tests-disabled Release guard and was not rebuilt in this focused Microsoft Drive lane.
- [x] Enabled the explicit destination lookup hint in `MoveOrRenameItem`; direct callers still fetch normally, while the `ENABLE_TESTS` legacy scope can force the old refetch for same-binary comparison.
- [x] Captured the first GREEN candidate run under `.build/TestSandbox/runs/20260721T104000Z-msdrive-merge-green/`. Microsoft Drive reported 121 passed / 0 failed. Baseline measured 83 logical metadata GETs / 84 total Graph GETs / 16 PATCHes / 1,322,934 us; candidate measured the exact target 67 / 68 / 16 / 1,058,244 us. The candidate removed one GET per child and improved fixed-latency elapsed time by 264,690 us (20.0%). The executable still reported only the same two unrelated stale S3 test-export failures; the final full test-enabled rebuild remains responsible for restoring those artifacts before evidence capture.
- [x] Added focused exact-count/identity contracts for existing-file overwrite, Retry-to-missing, nested folder merge, and a destination created after the first missing probe. The first focused run found that overwrite's collision-safe rollback-name probe was not included in the new logical metadata aggregate; production behavior was correct, but the measurement was incomplete. Routed that probe through `MoveMetadataMetrics` and locked the complete overwrite shape at 8 logical metadata GETs / 9 total GETs / 2 PATCHes / 1 DELETE.
- [x] The focused rerun under `.build/TestSandbox/runs/20260721T104400Z-msdrive-merge-adversarial-green/` passed all 137 Microsoft Drive checks. The new contracts prove Retry replaces the hint (8 logical / 9 total GETs), nested recursion retains listings while avoiding the leaf duplicate (8 / 10), and a late destination collides at PATCH without source loss or DELETE (7 / 8). The full executable still has only the known stale S3 test-export artifact failures pending the full test-enabled rebuild.
- [x] Added a permanent source contract for the `ENABLE_TESTS` helper/provider gate, Microsoft Drive's canonical Common project reference, explicit hint consumption and Retry replacement, rollback-probe metric coverage, exact 83-to-67 formulas, all focused edge tests, and all six metric names. The same edit widened an older Riptide call-site regex to accommodate the new optional metric/hint arguments without weakening its `checkCancel` requirement. `TestHarnessSourceContracts.Tests.ps1`: 149 passed / 0 failed.
- [x] Full test-enabled Release attempt 1 reached normal solution compilation with no warnings/errors shown, then the 120-second command wrapper expired before MSBuild completed. This was an inconclusive harness timeout, not a compiler/test failure.
- [x] The incremental retry was intentionally rejected by `.build/artifact-operation-contaminated.json`; the build guard required a same-configuration full-solution `-Rebuild` after interruption.
- [x] Full-solution test-enabled Release recovery attempt 2 compiled the provider/application graph through `RedSalamander.exe`, `RedConfigure.exe`, and most final test projects with no surfaced warning/error, but the 15-minute outer wrapper expired during `MonitorTest`. Attempt 3 below completed the required recovery.
- [x] Full-solution test-enabled Release recovery attempt 3 completed successfully with four workers in 9:38 (`.build/logs/msbuild-20260721_130315_812.log`), clearing the contamination marker. It reported 0 errors and 10 unrelated `ViewerPETests.cpp` C5245/C5264 unused test-helper warnings; the Microsoft Drive focused builds remain 0 warnings/errors.
- [x] The first unified-artifact Release contract run under `.build/TestSandbox/runs/20260721T131000Z-msdrive-merge-release-check/` passed completely: Microsoft Drive 137/137, S3 multipart 26/26, S3 directory-transfer probes 16/16, and `PluginContractTests` exit 0. Baseline/candidate were 83/84 GETs at 1,326,290 us versus 67/68 at 1,055,017 us (271,273 us / 20.45% faster).
- [x] Archived five same-binary Release samples under `Specs/TestRuns/7d3a1247382a/FileOps/2026-07-21_131329_msdrive_merge_metadata/`. Every run kept the exact 83/84-to-67/68 request shape and passed the complete executable. Median fixed-latency time improved from 1,323,954 us to 1,059,139 us: 264,815 us / 20.00%. The 12-file, 386,436-byte curated archive parses and passes explicit `TestRunArchive` size/profile validation with zero violations; no p95 claim is made from five operation-level samples.
- [x] Full Debug solution build passed in 6:01 with 0 warnings/errors (`.build/logs/msbuild-20260721_131613_290.log`). Debug `PluginContractTests` then exited 0: Microsoft Drive 137/137, FileSystem 52/52, 7z 26/26, Google Drive 9/9, S3 150/150, Curl 78/78, S3 multipart 26/26, and S3 directory probes 16/16. Its trace is curated as `debug-contract-trace.txt` in the evidence archive.
- [x] Production tests-disabled x64 Release Microsoft Drive and Common rebuilt with 0 warnings/errors (`.build/logs/msbuild-20260721_132254_234.log`). `dumpbin /exports` reports zero `RedSalamanderMicrosoftDriveDebugSelfTests` exports, proving the Release-capable fake Graph/selftests remain absent from the shipped DLL.
- [x] Passed focused provider/source contracts and Debug, test-enabled Release, and production tests-disabled Release build gates. The broad Release rebuild's 10 warnings are unrelated pre-existing ViewerPETests unused-helper warnings; all focused Microsoft Drive and Debug builds are warning-free.
- [x] Archived same-machine Release evidence and deferred `$batch`/pacing until evidence justifies a separate semantics-changing design.
- [x] Merged the lasting contract into `Specs/FileSystem/FileSystem_MicrosoftDrive.md` and `Specs/Testing/Testing_TestCoverage.md`, reconciled FG-P2-6, and completed the move-to-Done gate.

## Done Gate

Satisfied. The candidate removes exactly one destination metadata GET per ordinary merge child; deterministic tests preserve missing, overwrite, Retry, nested, and concurrent-destination behavior; five same-machine test-enabled Release samples prove the exact request-count reduction and 20.00% median fixed-latency improvement; Debug, test-enabled Release, production Release, provider, source-contract, archive, and diff gates are green; the authoritative Microsoft Drive/testing specs own the lasting contract; and FG-P2-6 records the final batching decision.
