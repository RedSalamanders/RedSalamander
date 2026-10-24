# Operation FileOperations Fault-Injection Ratchet - 2026-07-05

Owner for the long-term remainder routed from Granite GR-A5 after the first
interface-boundary bridge IO decorator slice landed.

## Status

**DONE 2026-07-20.** All remaining FileOperations bridge fault hooks now use
the role-aware I/O decorator or have a reviewed, source-contract-locked
call-site exception. FIR-1 has serial and parallel behavioral proof; FIR-4
replaces the obsolete broad occurrence counts with an exact structural
allowlist. The authoritative bridge contract contains the durable behavior.

## 2026-07-20 Reconciliation - Last Finding and Current Code State

The destination-side seam requested by FIR-1/FG-A1 is already implemented:

- `SelfTestBridgeFileReader::GetSize` consumes
  `g_fileOpsBridgeFailNextDestinationGetSizeCount` only for
  `SelfTestBridgeIoRole::Destination` and returns `ERROR_READ_FAULT`.
- the bridge decorates `destinationFileSystemIo` with the Destination role;
- queue/header setters and attempt counters exist, and the common FileOps
  self-test cleanup resets them.

The missing piece at reconciliation time was proof: no self-test armed
`SetFileOpsBridgeFailNextDestinationGetSizeForSelfTest`, so the post-promote
destination re-stat branch had no deterministic behavioral coverage. The
closeout slice added serial and concurrency-greater-than-one MOVE cases,
asserted `ERROR_PARTIAL_COPY`, destination-byte integrity, and source
preservation, then locked registration and seam wiring in
`TestHarnessSourceContracts.Tests.ps1`.

Progress ledger (persisted during implementation):

| Date | State | Finding / action | Evidence |
|------|-------|------------------|----------|
| 2026-07-20 | DONE | Reconciled plans with code; found the Destination decorator seam and public self-test controls already present but unused by any behavioral test. | `FolderWindow.FileOperations.State.cpp`, `.State.Queue.cpp`, `FileOperationsInternal.h`, and repository search for the setter. |
| 2026-07-20 | DONE | Added a serial Destination-role post-promote failure case and an eight-child concurrency-four MOVE case that injects one source and one destination `GetSize` failure. Added FIR-2/FIR-3 code comments, authoritative contract text, family registration, and the FIR-4 structural allowlist. | Debug x64 builds `msbuild-20260720_204342_943.log` and `msbuild-20260720_205005_819.log`: 0 warnings, 0 errors. |
| 2026-07-20 | DONE | First focused run proved the parallel copy-phase partial result suppresses the entire tree's MOVE cleanup pass: all 8 sources remained, 7 valid destinations existed, with one source-only and seven source+destination pairs. Corrected the test and authoritative contract from the narrower per-unverified-file expectation. | Discovery run `20260720T184827Z-52768-70d949890eb24cbf988530ef5e595b10`; the corrected focused run archived at `Specs/TestRuns/7d3a1247382a/FileOps/2026-07-20_205413/`. |
| 2026-07-20 | DONE | Revalidated the focused prefix, full family, and structural suite after the finding. Repaired one stale, unrelated Track 14 textual-order assertion so the required source-contract file is wholly green against current code. | `Floodgate_CrossFsMove*`: 6/6; `FileOpsFamily_Fairstream`: 47/47, archive `Specs/TestRuns/7d3a1247382a/FileOps/2026-07-20_205550/`; `TestHarnessSourceContracts.Tests.ps1`: 148/148. |

## Baseline

Historical baseline at routing time:

- `RedSalamander/FolderWindow.FileOperations.State.cpp`: 31 `ForSelfTest`
  occurrences.
- `RedSalamander/FolderWindow.FileOperations.State.cpp`: 30
  `#ifdef ENABLE_TESTS` blocks.
- `RedSalamander/`: 629 `ForSelfTest` occurrences.
- `RedSalamander/`: 647 `#ifdef ENABLE_TESTS` blocks.

The raw repository counts on 2026-07-20 are 59/39 in
`FolderWindow.FileOperations.State.cpp` and 930/693 under `RedSalamander/`
(`ForSelfTest` / `#ifdef ENABLE_TESTS`). Those totals increased because later
plans added broad deterministic self-test coverage; they no longer distinguish
approved boundary decorators from pipeline coupling. FIR-4 therefore replaces
the broad occurrence-count ratchet with a scoped source-contract allowlist of
the remaining bridge call-site injections. New bridge data-safety failures must
still use the decorator seam unless this plan and the allowlist document the
boundary mismatch.

The first Granite GR-A5 ratchet slice established
`SelfTestBridgeIoDecorator` / `SelfTestBridgeFileReader`, injected through
`DecorateBridgeIoForSelfTest(fileSystemIo, SelfTestBridgeIoRole::Source)`, and
migrated the cross-filesystem bridge source `GetSize` fail-next hook out of the
inline `CopyFileWithBuffer(...)` consumption path.

## Rules

- New FileOperations bridge data-safety fault-injection tests should use the
  interface-boundary decorator seam when the failure belongs to `IFileSystemIO`
  or `IFileReader`.
- Do not add new inline production-pipeline globals/call-site hooks in
  `FolderWindow.FileOperations.State.cpp` unless the hook cannot be expressed at
  an existing boundary; document any exception in this plan.
- When a touched pipeline site still has an inline self-test hook, migrate it to
  the decorator seam or a shared helper before adding adjacent coverage.
- Treat the historical broad counts above as context only. The FIR-4 structural
  allowlist is the authoritative ratchet because it distinguishes reviewed
  boundary mismatches from ordinary test-only coverage growth.

## Ratchet Slice Disposition

| ID | Status | Slice | Direction |
|----|--------|-------|-----------|
| FIR-1 | DONE | Destination-side bridge `GetSize` fault injection, tracked as Floodgate FG-A1. | Destination-role decorator seam plus serial post-promote and concurrency-four source/destination failure coverage are implemented, family-reachable, and green. |
| FIR-2 | APPROVED EXCEPTION | Bridge file-copy failure hook is consumed at the `CopyFileWithBuffer` operation boundary. | It injects a whole-copy worker failure before stream I/O to prove exact worker-HRESULT propagation and paused-reader shutdown in both serial and parallel scheduling. It is not an `IFileReader`/`IFileWriter` contract failure, so moving it into a provider decorator would change the tested failure layer. FIR-4 locks the single named call site. |
| FIR-3 | APPROVED EXCEPTIONS | Destination mutation and create-directory race helpers remain at two pipeline transition points. | The race must occur after the absent-path probe and before `CreateDirectory`; the mutation must occur after promote/manifest record and before MOVE cleanup verification. An I/O decorator cannot infer either temporal boundary. Both helpers mutate through the real destination interfaces, are deterministic, and FIR-4 locks their single named call sites. |
| FIR-4 | DONE | Replace obsolete broad counts with a structural source-contract allowlist. | The contract asserts Destination-role wiring/counter behavior, serial/parallel family reachability, and exactly one call site for every FIR-2/FIR-3 exception. |

## Closeout Gate

Satisfied 2026-07-20:

- I/O-contract faults use the boundary decorator; all three remaining call-site
  exceptions have explicit boundary-mismatch rationales and an exact allowlist.
- focused `Floodgate_CrossFsMove*` passed 6/6 and the full Fairstream family
  passed 47/47 with no skips;
- `TestHarnessSourceContracts.Tests.ps1` passed 148/148;
- this slice changes only `ENABLE_TESTS` behavior and documentation, so it makes
  no production throughput/latency claim requiring a baseline/candidate perf
  comparison. The FileOps runs are nevertheless archived with emitted metrics.
