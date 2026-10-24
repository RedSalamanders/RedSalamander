# S3 Interactive Directory Destination-Probe Optimization

**Status:** Done  
**Date:** 2026-07-21  
**Parent:** `Specs/Plans/WIP/Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md` (`FG-P2-5`)  
**Protected scenario:** interactive S3 directory COPY/MOVE of many objects into a destination prefix, with no conflicts on the normal path

## Objective

Remove the redundant whole-plan destination metadata preflight from S3 directory transfers when the operation can resolve conflicts interactively or overwrite is already authorized. Preserve the just-in-time per-object refresh, Retry re-probe, ancestor-blocker removal re-probe, and the no-callback/no-overwrite fail-closed preflight.

## Last Finding and Current Code State

The original finding was an unconditional whole-plan `RefreshDestinationState` pass whose results were overwritten by the just-in-time per-object refresh on the interactive path. Because one refresh serially probes possible object-as-directory ancestors and the leaf, this doubled the protected scenario's destination metadata traffic. The last implementation review confirmed the safe boundary is to remove only that behaviorally unused pass; reusing cached preflight state would widen the race window and could silently overwrite a destination created after planning.

`ExecuteCopyOrMove` now computes `requiresFailClosedDestinationPreflight`. The whole-plan pass runs only when overwrite is not authorized and no `reportIssue` callback exists, retaining the before-progress/before-mutation `ERROR_ALREADY_EXISTS` contract. Interactive and pre-authorized-overwrite transfers begin with the existing just-in-time per-object refresh. Retry and successful ancestor-blocker removal still `continue` the decision loop and therefore obtain fresh state. Planned destination ancestors remain protected from deletion.

Aggregate refresh/probe count and probe-time instrumentation is present in the production path. A test-enabled S3 fake graph supplies exact request counts and fixed metadata latency, while a test-only scope restores the legacy interactive pass for same-binary comparison. The Release-capable provider contract is exported through `PluginContractTests`; structural source contracts lock the conditional, the just-in-time loop, exact Retry/ancestor behavior, metrics, and test wiring. The test-only graph and export are absent from the normal tests-disabled Release DLL.

The current working tree also contains the completed, uncommitted bridge commit-size proof optimization and unrelated user-owned WIP-plan edits. FG-P2-5 changes remain confined to the S3 directory implementation/test export, provider-test and source-contract wiring, S3/Testing specifications, this plan and its parent, and the dedicated evidence archive.

## Required Contract

- Interactive transfers perform no whole-plan destination preflight; each object receives one just-in-time destination refresh on the no-conflict first attempt.
- Transfers with `FILESYSTEM_FLAG_ALLOW_OVERWRITE` perform no fail-closed preflight even without a callback.
- No-callback transfers without overwrite permission retain the complete preflight and return `ERROR_ALREADY_EXISTS` before progress or mutation when any planned destination is blocked.
- A Retry answer re-probes destination and ancestor state before making another decision.
- Removing an ancestor blocker re-probes because deeper stacked blockers may remain.
- Planned destination ancestor collisions keep their data-safety behavior: the descendant is skipped rather than deleting a sibling created by the same transfer.
- Production metrics are aggregate per transfer and do not emit one JSONL row per object or per HEAD request.

## Measurement and Deterministic Proof

The test-enabled S3 fake graph will count summary probes and optionally add fixed deterministic metadata latency. A many-object no-conflict interactive scenario will execute the legacy double-pass baseline and optimized candidate against identical graphs in the same Release binary.

Planned metric family:

- `FileOps.S3.DirectoryTransfer.DestinationProbeCount`
- `FileOps.S3.DirectoryTransfer.DestinationProbeUs`
- `FileOps.S3.DirectoryTransfer.DestinationRefreshCount`
- `FileOps.S3.DirectoryTransfer.Baseline`
- `FileOps.S3.DirectoryTransfer.Candidate`
- `FileOps.S3.DirectoryTransfer.Improvement`

The exact leaf/ancestor HEAD count depends on destination depth. The invariant is one complete refresh per object on the normal candidate path versus two in the legacy baseline. The deterministic scenario must assert exact refresh and underlying summary-probe counts, destination bytes, source preservation for COPY, and no unexpected prompts.

## Progress Log

- [x] Re-verified the finding against current code: unconditional preflight at `FileSystemS3.Directory.cpp:1530-1545`, unconditional per-object refresh at `1585-1592`, Retry `continue`, and ancestor-removal `continue` are all still present.
- [x] Corrected the implementation direction: do not reuse potentially stale preflight state on the interactive path; skip the behaviorally unused preflight and retain just-in-time per-object probing.
- [x] Recorded the existing dirty-worktree boundary and the full Done gate before code changes.
- [x] Added per-transfer aggregate destination refresh/probe count and probe-time instrumentation. The production path emits three coalesced `FileOps.S3.DirectoryTransfer.*` rows only when perf capture is active.
- [x] Extended the S3 fake graph under `ENABLE_TESTS` with exact summary-probe counts and fixed latency, added a test-only legacy-preflight scope, and wired a Release-capable provider export through `PluginContractTests`.
- [x] Added the deterministic 16-object baseline/candidate scenario plus exact no-callback fail-closed, Retry, and stacked-ancestor re-probe assertions. The production preflight is intentionally still unconditional at this checkpoint, so the candidate assertions are expected to be RED until the optimization lands.
- [x] Test-enabled full Release rebuild completed successfully after recovering from an outer-command timeout with the repository-required `-Rebuild -MaxCpuCount 4`. Log: `.build/logs/msbuild-20260721_105139_426.log`; 0 errors. Its 10 warnings are pre-existing `C5245`/`C5264` unused symbols in `Tests/ViewerPETests/ViewerPETests.cpp`; the changed S3 and PluginContractTests projects emitted no warning.
- [x] RED Release run captured under `.build/TestSandbox/runs/20260721T110300Z-s3-directory-probe-red/`: provider slice reported 10 passed / 4 failed. Legacy and not-yet-optimized candidate both measured 32 refreshes, 64 destination probes, and 1,009,303 / 1,008,938 us fixed-latency probe time. Retry measured 3 refreshes / 6 probes and stacked ancestors 4 / 14, proving each carried the same redundant preflight. Byte correctness and the no-callback fail-closed assertions passed. Curated baseline archive is pending the final paired evidence package.
- [x] Archived five same-binary Release comparisons at `Specs/TestRuns/7d3a1247382a/FileOps/2026-07-21_111500_s3_directory_probe/`. All five `PluginContractTests` executions exited 0; the S3 directory slice passed 16/16 every time. Legacy shape was invariant at 32 refreshes / 64 probes and candidate at 16 / 32. Median fixed-latency probe time improved 1,011,056 us -> 504,449 us (-50.11%). The archive is 493,366 bytes total, largest file 151,230 bytes, and all 260 JSONL rows match profile `7d3a1247382a`.
- [x] `Test-TestRunArchive.ps1` was attempted and its only failures are inherited from the preceding uncommitted archive `Specs/TestRuns/7d3a1247382a/FileOps/2026-07-20_205550/` (6,991,810-byte JSONL; 7,050,178-byte run), not this plan's archive. FG-P2-5's archive independently passes the validator's 2 MiB/file and 5 MiB/run limits plus machine-profile/schema checks. The prior evidence is left untouched pending its owning task's reconciliation.
- [x] Implemented `requiresFailClosedDestinationPreflight`: production preflight now runs only when overwrite is not authorized and no conflict callback exists. The test-only legacy scope can restore the old interactive pass for same-binary comparison. The just-in-time loop, Retry `continue`, and ancestor-removal `continue` are unchanged.
- [x] Added a focused no-callback + `FILESYSTEM_FLAG_ALLOW_OVERWRITE` assertion proving that already-authorized replacement also skips preflight, performs one refresh / two leaf-depth probes, replaces destination bytes, and preserves the COPY source.
- [x] Focused test-enabled Release S3 rebuild passed with 0 warnings/errors (`.build/logs/msbuild-20260721_110538_535.log`). First GREEN provider run under `.build/TestSandbox/runs/20260721T111000Z-s3-directory-probe-green/` passed the new slice 16/16 and the complete `PluginContractTests` executable. Same-binary metrics: forced legacy 32 refreshes / 64 probes / 1,006,287 us; candidate 16 / 32 / 507,503 us; fixed-latency reduction 49.6%. Retry is 2 / 4, stacked ancestors 3 / 11, and pre-authorized overwrite 1 / 2.
- [x] Debug `FileSystemS3` and `PluginContractTests` focused builds passed with 0 warnings/errors (`msbuild-20260721_111105_618.log`, `msbuild-20260721_111241_308.log`). Debug provider contracts passed existing S3 150/150, multipart 26/26, and directory probes 16/16; the complete executable including unload quiet points and shared-libcurl refresh proof exited 0.
- [x] Source-contract suite passed 148/148 after locking the conditional preflight, just-in-time refresh, exact scenario counts, metric family, Release export, provider-test wiring, and metric-aware final partial-copy result; the final post-spec-reconciliation rerun also passed 148/148.
- [x] Full Debug solution build completed cleanly with refreshed `FileSystemS3.dll` and `PluginContractTests.exe`: `.build/logs/msbuild-20260721_111508_805.log`, 0 error matches and 0 warning matches.
- [x] Production tests-disabled Release `FileSystemS3` build passed with 0 warnings/errors (`.build/logs/msbuild-20260721_112521_940.log`), proving the fake graph, legacy comparison switch, and self-test export remain correctly guarded from the normal Release DLL.
- [x] Passed focused S3 provider contracts, source contracts, clean full Debug, test-enabled Release, and production tests-disabled Release build lanes.
- [x] Merged the durable preflight/freshness/re-probe contract, metric definitions, deterministic proof scope, evidence link, and live-S3 caveat into `Specs/FileSystem/FileSystem_S3.md`.
- [x] Reconciled FG-P2-5 as Done in the parent plan, updated `Specs/Testing/Testing_TestCoverage.md`, completed a clean final diff review, and moved this plan to `Specs/Plans/Done/`.

## Done Gate

This plan may move to `Done` only when the optimized path is code-complete; deterministic tests prove exact normal, fail-closed, Retry, and ancestor-removal probe behavior without changing transfer results; same-machine test-enabled Release evidence is archived and demonstrates one refresh per object plus a material fixed-latency improvement; Debug and Release builds and provider/source-contract checks are green; the authoritative S3 spec owns the lasting behavior and metric contract; and the parent FG-P2-5 row is marked complete.
