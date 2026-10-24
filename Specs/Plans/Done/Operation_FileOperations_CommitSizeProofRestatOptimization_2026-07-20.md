# File Operations Commit-Size Proof Re-stat Optimization

**Status:** Done - implementation, durable contracts, performance proof, and regressions complete  
**Date:** 2026-07-20  
**Parent:** `Specs/Plans/WIP/Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md` (`FG-P0-2-PERF-RESTAT`)  
**Protected scenario:** cross-filesystem MOVE of many small files to a destination with expensive metadata reads

## Objective

Remove the redundant immediate final-path destination `CreateFileReader` + `GetSize` round-trip after a successful bridge MOVE without weakening the source-deletion gate. COPY keeps its post-commit destination reader because it is reused for the required destination-content hash. MOVE keeps the later cleanup-time destination size and hash verification before source deletion.

## Last Finding and Pre-change State

The bridge performed two destination metadata reads for every known-size moved file: one immediately after writer `Commit`/promotion and another during MOVE cleanup. Only the cleanup-time read participates in the final source-deletion proof. `IFileWriter::GetPosition` cannot replace the first read because it reports stream position, not successfully committed object size. Provider identity, atomic-writer capability, and staging-name shape are also not integrity proofs.

The fail-closed boundary is therefore an optional COM interface on the concrete writer. A writer may report an exact, writer-local committed byte count only after `Commit` succeeds. The host accepts it only when it exactly matches the known source size. Missing, failed, or mismatched proof retains the legacy final-path re-stat. Successful staged promotion preserves the proof because it renames/moves the already committed bytes; it does not authorize source deletion.

## Required Contract

- `IFileWriterCommitSizeProof` has its own UUID and exposes `GetCommittedSize(uint64_t*)`.
- Pre-Commit calls fail with `HRESULT_FROM_WIN32(ERROR_INVALID_STATE)`; null output fails with `E_POINTER`.
- A successful call is writer-local and performs no filesystem, service, or network request.
- Proof-capable MOVE skips only the immediate destination size probe.
- COPY still opens and hashes the final destination.
- MOVE cleanup still opens the final destination and verifies exact size plus content hash before deleting the source.
- Writers without a valid proof follow the unchanged legacy final-path verification path.
- Host logic is capability-driven through `QueryInterface`; it must not branch on provider IDs.

## Performance Contract and Metrics

The deterministic Fairstream case `Floodgate_CrossFsMoveCommitSizeProofAvoidsImmediateRestat` moves 16 local 256-byte files to Dummy with 25 ms metadata latency and concurrency four. It verifies every destination byte and source removal. Test-enabled builds may set `REDSALAMANDER_FILEOPS_BRIDGE_DISABLE_COMMIT_SIZE_PROOF=1` to produce the same-binary fallback baseline; normal execution requires the candidate shape.

Per-operation metrics:

- `FileOps.Bridge.ImmediateDestinationSizeProbeCount`
- `FileOps.Bridge.ImmediateDestinationSizeProbeUs`
- `FileOps.Bridge.MoveCleanupDestinationSizeProbeCount`
- `FileOps.Bridge.MoveCleanupDestinationSizeProbeUs`
- `FileOps.Bridge.CommitSizeProofCount`
- `FileOps.Bridge.CommitSizeProofFallbackCount`

Expected 16-file shapes:

| Shape | Proof | Fallback | Immediate probes | Cleanup probes |
|---|---:|---:|---:|---:|
| Same-binary fallback baseline | 0 | 16 | 16 | 16 |
| Candidate | 16 | 0 | 0 | 16 |

## Progress Log

- [x] Instrumented immediate and cleanup destination size-probe counts/timing plus proof/fallback counts.
- [x] Captured initial test-enabled Release baseline at `Specs/TestRuns/7d3a1247382a/FileOps/2026-07-20_212924/`: 16 immediate probes / 506,623 us; 16 cleanup probes / 490,527 us; total protected probe time 997,150 us.
- [x] Added `IFileWriterCommitSizeProof` and implementations for Dummy, S3 multipart, and Microsoft Drive writers.
- [x] Updated the bridge to accept only exact post-Commit proof, retain fail-closed fallback, and skip the immediate re-stat only for MOVE.
- [x] Added provider pre/post-Commit contracts, legacy destination-`GetSize` fault assertions, the 16-file behavioral/perf case, and source-contract guards.
- [x] First candidate Release run archived at `Specs/TestRuns/7d3a1247382a/FileOps/2026-07-20_213832/`: 0 immediate probes; 16 cleanup probes / 489,140 us; total protected probe time improved 50.9% versus the initial baseline. Focused test passed 3/3, `Floodgate_CrossFsMove*` passed 7/7, source contracts passed 148/148, and Release `PluginContractTests.exe` passed.
- [x] Added a test-only proof-disable switch, restricted proof queries/counters to MOVE, and rebuilt the finalized test-enabled Release binary cleanly (0 warnings/errors). That same binary passed 15/15 checks for five fallback repeats and 15/15 for five candidate repeats. Final archives: fallback `Specs/TestRuns/7d3a1247382a/FileOps/2026-07-20_215537/`; candidate `Specs/TestRuns/7d3a1247382a/FileOps/2026-07-20_220125/`.
- [x] Final quality-gated analysis passed with five samples required for p95 and p99: protected p95 improved from 1,000,350 us to 490,890 us (-50.9%), with 0 regressions, 1 improvement, and 0 noise. Every fallback sample enforced 16 immediate + 16 cleanup probes; every candidate sample enforced 0 immediate + 16 cleanup probes.
- [x] Updated `Specs/Core/Core_FileSystemBridge.md` and `Specs/Plugins/Plugins_VirtualFileSystem.md` with the durable host/plugin contract, fallback behavior, metrics, provider support, and source-deletion safety boundary.
- [x] Source-contract suite passes 148/148 after locking MOVE-only query/fallback behavior, the test-only comparison hook, all provider implementations, and both metric shapes.
- [x] Finalized Release `PluginContractTests.exe` passes, including S3 multipart proof selftests (`passed=26, failed=0`). A Release Fairstream-family attempt stopped after 3 passes at unrelated `Riptide_ReparseMoveRollbackKeepsOverwrittenDestination`; focused repetition reproduced the missing injection 5/5. Diagnosis: that local-plugin hook and its environment constants are gated by `_DEBUG` in `Plugins/FileSystem/FileSystem.FileOps.cpp`, so the test-enabled Release host loads the plugin's intentional non-injecting stub. This is a pre-existing configuration mismatch, not a production MOVE failure or a commit-size-proof path. Run `20260720T200209Z-55536-9acce90fae7b4333a1adbf98fb332acd` retains the evidence. The authoritative Debug Fairstream-family run below contains the required provider hooks and is green.
- [x] Debug build passes with 0 warnings/errors. Debug `PluginContractTests.exe` passes, including Microsoft Drive `passed=115, failed=0`, S3 `passed=150, failed=0`, and S3 multipart `passed=26, failed=0`; the new pre/post-Commit proof checks execute in these provider suites.
- [x] The full Debug Fairstream family passes 48/48 with no failures, skips, flakes, or classification issues (run `20260720T200857Z-54872-93562e71f37a43d9a60dc0c65244f28f`).
- [x] A repository-wide Debug `-Suite Full -SkipBuild` attempt reached the 30-minute outer command ceiling before producing a consolidated result. Its partial Commands artifact completed 348 passes / 11 failures, all in pre-existing settings, Connection Manager, or Preferences UI cases; Compare did not finish and FileOps had not started. No failure touches `IFileWriter`, provider proof implementations, or the bridge, so this attempt is classified as inconclusive/unrelated rather than proof for or against the change. The interrupted-operation safeguard was honored with a full-solution Debug rebuild (0 warnings/errors), after which complete Debug FileOps passed 111/111 runnable cases with 0 failures and 20 expected remote/configuration-gated skips (run `20260720T204859Z-49168-6638d16156eb4c4684543e291b295826`).
- [x] Final repeated before/after evidence is archived and `Show-PerfRuns.ps1 -FailOnQuality` passes with explicit five-sample p95/p99 thresholds.
- [x] Parent row is reconciled, authoritative specs own all durable behavior, and this plan is moved to `Specs/Plans/Done/`.

## Done Gate

This plan may move to `Done` only when the Release and Debug builds are clean; focused, complete FileOps, family, source-contract, and provider-contract checks are green; repeated same-machine archived evidence proves 0 immediate and exactly 16 cleanup probes for the candidate; legacy/fallback evidence proves 16 immediate and 16 cleanup probes; authoritative specs describe the interface and safety boundary; unrelated repository-wide failures are explicitly classified; and the parent Floodgate row is marked complete.
