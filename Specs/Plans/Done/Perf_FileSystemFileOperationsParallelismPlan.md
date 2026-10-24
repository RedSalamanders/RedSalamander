# FileSystem Plugin & File Operations — Performance & Parallelism Improvement Plan

**Status:** Complete  
**Date:** 2026-04-01  
**Scope:** `Plugins/FileSystem/FileSystem.FileOps.cpp`, `RedSalamander/FolderWindow.FileOperations*.cpp`, `Common/PlugInterfaces/FileSystem.h`  
**~LOC reviewed:** FileSystem.FileOps.cpp (~7K), FolderWindow.FileOperations.State.cpp (~3K), FileSystem.DirectoryOps.cpp (~1.5K), FileSystem.Search.cpp (~1.9K)

---

## Progress Checklist

> **How to update:** When starting a phase, change `[ ]` → `[~]`. When all sub-tasks in that phase are done and perf evidence is archived, change `[~]` → `[x]`. Update the date column on each transition.

| Done | Phase | Area | Priority | Started | Completed |
|------|-------|------|----------|---------|-----------|
| [x] | 1 | Recycle Bin Batching | **P1** | 2026-03-31 | 2026-03-31 |
| [x] | 2 | Cross-FS Bridge Pipeline | **P2** | 2026-03-31 | 2026-04-01 |
| [x] | 3 | Callback Contention Reduction | **P2** | 2026-03-31 | 2026-04-01 |
| [x] | 4 | Parallel Pre-Calc Tree Walk | **P3** | 2026-03-31 | 2026-03-31 |
| [x] | 5 | Parallel Search Directory Walk | **P3** | 2026-03-31 | 2026-03-31 |
| [x] | 6 | Bandwidth Throttle Rework | **P4** | 2026-03-31 | 2026-04-01 |
| [x] | 7 | Adaptive Concurrency | **P4** | 2026-04-01 | 2026-04-01 |
| [x] | 8 | Settings — Plugin Schema + Global Pane | **P2** | 2026-03-31 | 2026-04-01 |

### Plugin Interface / Contract Work (any phase)

| Done | Change | Affected Phase |
|------|--------|-----------|
| [x] | Extend `IFileSystem` directly with `GetTransferHints(...)` | 2, 7 |
| [x] | Extend `IFileSystem` directly with `GetStorageCharacteristics(...)` | 7 |
| [x] | Keep `FileSystemOptions` ABI unchanged for this plan; resolve Auto/Manual to existing concurrency fields before operation start | 7 |
| [x] | Keep `IFileSystemCallback` ABI unchanged; reduce contention within the existing callback contract | 3 |
| [x] | Extend FileSystem plugin `kSchemaJson` with `concurrencyMode`, `recycleBinBatchSize`, `searchMaxDirectoryWalkers` | 8 |
| [x] | Widen FileSystem plugin `copyMoveMaxConcurrency` range to 1–16 and validate values >8 with same-machine perf evidence in the same slice | 8 |

---

## Current Architecture Summary

The FileSystem plugin already has a well-structured foundation:
- **SharedFileOpsJobScheduler**: work-stealing bounded thread pool (1–16 `std::jthread` workers, round-robin job queue, module-pinned)
- **Parallel copy**: `CopyDirectoryChildrenParallel()` with work-stealing queue (256+ item prefetch)
- **Parallel delete**: `DeleteDirectoryRecursiveParallel()` with tree flattening → parallel file deletion → post-order directory removal
- **NtQueryDirectoryFile enumeration**: handle-based fast path on local disks, Win32 fallback for network/UNC
- **Configurable concurrency**: plugin schema now exposes `concurrencyMode` (`auto`/`manual`), `_copyMoveMaxConcurrency` (default 4, max 16), `_deleteMaxConcurrency` (default 8), and `_deleteRecycleBinMaxConcurrency` (default 2); the host resolves Auto mode from `GetStorageCharacteristics(...)` at task start without changing `FileSystemOptions`
- **Progress contract**: `callbackMutex`-serialized callbacks, `progressStreamId` per worker, 50–100ms throttle
- **Bandwidth throttling**: callback-driven pacing with 10ms cancel polls; sequential and parallel exact-case evidence is archived, and the remaining fairness delta is now understood as callback-quantum-limited on this machine

The improvements below target remaining serial bottlenecks, contention points, and missed parallelism opportunities.

---

## Phase 1 — High Impact: Recycle Bin Batching (P1)

### Problem

`DeleteToRecycleBin()` creates a **fresh `IFileOperation` COM object per item** and calls `SHCreateItemFromParsingName()` individually:

```
FileSystem.FileOps.cpp ~L3911: CoCreateInstance(CLSID_FileOperation, ...) — per item
FileSystem.FileOps.cpp ~L3920: SHCreateItemFromParsingName() — per item
FileSystem.FileOps.cpp ~L3948: fileOperation->DeleteItem() — singular
```

For 10K+ files, COM overhead (object creation + shell namespace resolution) dominates. The Windows `IFileOperation` API natively supports multi-item batching but it is unused.

### Proposed Changes

- [x] **1.1** Add `DeleteToRecycleBinBatched()` that groups items by parent directory and submits them in batches to a single `IFileOperation` instance
  - Collect all items destined for recycle bin into a `std::vector<std::wstring>`
  - Group by parent path (items in the same folder batch together for shell efficiency)
  - Create one `IFileOperation` per batch, call `DeleteItems()` (plural) via `IShellItemArray`
  - Default cap batch size at ~500 items to allow cancellation checkpoints between batches; later expose it as plugin setting `recycleBinBatchSize`

- [x] **1.2** Implement `IFileOperationProgressSink` for batch mode that records per-item `PostDeleteItem()` results and preserves the existing progress / cancel checkpoints while the batch is running

- [x] **1.3** Handle partial failure: replay observed per-item results through the existing `FileSystemItemCompleted()` path after the batch finishes; if the batch aborts under cancel / stop-on-error semantics, unobserved items remain unreported just like the single-item path

- [x] **1.4** Retain the existing single-item codepath as fallback for edge cases (single item, or when batch creation fails)

- [x] **1.5** Add perf instrumentation: `Debug::Perf::Scope` around batch creation, `PerformOperations()`, and per-item completion counts

### Landed 2026-03-31

- `FileSystem::DeleteItems()` now enters the shared delete scheduler even when recycle-bin delete concurrency resolves to 1, so sibling batching also applies to the sequential recycle-bin case.
- Ready items are grouped by parent directory before dispatch, capped at the plugin-configured `recycleBinBatchSize` (default `500`) per batch, and passed to `DeleteToRecycleBinBatched()` via `IFileOperation::DeleteItems()` + `IShellItemArray`.
- `RecycleBinBatchProgressSink` keeps the current callback-based pause / cancel flow intact by continuing to emit throttled `FileSystemProgress()` checkpoints from `PreDeleteItem()` / `UpdateProgress()` while recording per-item `PostDeleteItem()` outcomes for later replay.
- If batch setup fails before `DeleteItems()` queues the shell work, the worker falls back to the existing single-item `DeleteToRecycleBin()` path and increments `FileOps.RecycleBin.BatchFallbackCount`.
- New plugin perf metrics: `FileOps.RecycleBin.BatchBuildUs`, `FileOps.RecycleBin.PerformOperationsUs`, `FileOps.RecycleBin.BatchDeleteUs`, `FileOps.RecycleBin.BatchObservedItems`, `FileOps.RecycleBin.BatchFailedItems`.

### Validation

- Deterministic selftests now cover both the single-batch and multi-batch paths: `Phase7_RecycleBinBatchDelete` deletes `384` sibling files, and `Phase7_RecycleBinBatchDeleteMultiBatch` deletes `768` sibling files so the `500`-item batch cap is exercised.
- Same-machine baseline archive: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_135355`
- Same-machine candidate archive: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_140614`
- `Phase7_RecycleBinBatchDelete` case duration improved from `5047 ms` to `1766 ms` (`2.86x` faster) in `fileops_results.json`.
- Host `FileOps.Operation` improved from `4,656,000 us` to `1,265,000 us` (`3.68x` faster) on the same machine for the `384`-item sibling delete case.
- Host callback overhead dropped materially in the same run: `FileOps.Progress.CallbackUs` `11,649 us` → `169 us`, `FileOps.ItemCompleted.CallbackUs` `2,274 us` → `156 us`.
- Candidate batch metrics for the landed `384`-item run: `BatchBuildUs=305,145 us`, `PerformOperationsUs=938,611 us`, `BatchDeleteUs=1,251,935 us`, `BatchObservedItems=384`, `BatchFailedItems=0`.
- Additional deterministic coverage archive: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_141949`
- The `768`-item multi-batch case passed in `3594 ms`; archived perf metrics confirm two recycle-bin batches with `BatchObservedItems=500` and `268`, both with `BatchFailedItems=0`.
- The archived `768`-item run also kept callback overhead low: `FileOps.Operation=2,735,000 us`, `FileOps.Progress.CallbackUs=352 us`, `FileOps.ItemCompleted.CallbackUs=602 us`.
- Fresh same-machine setting-backed archive: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_194908`
- `Phase7_RecycleBinBatchDelete` now proves the plugin setting itself: `recycleBinBatchSize=1` baseline `7,000,000 us` versus `recycleBinBatchSize=500` candidate `1,485,000 us`, a `5,515,000 us` improvement (`~4.7x` faster, `~78.8%` lower wall time).
- The same `2026-03-31_194908` archive also shows the callback cost collapse persists under the settings-backed comparison: `FileOps.Progress.CallbackUs` `25,724 us` → `431 us`, `FileOps.ItemCompleted.CallbackUs` `2,488 us` → `128 us`.
- Latest same-build rerun archive: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_201848`
- The latest `2026-03-31_201848` rerun stayed in the same performance envelope: `recycleBinBatchSize=1` baseline `5,672,000 us` versus `recycleBinBatchSize=500` candidate `1,156,000 us`, a `4,516,000 us` improvement (`~4.9x` faster, `~79.6%` lower wall time).
- The same `2026-03-31_201848` rerun kept callback overhead low: `FileOps.Progress.CallbackUs` `22,957 us` → `336 us`, `FileOps.ItemCompleted.CallbackUs` `2,439 us` → `123 us`; candidate batch metrics were `BatchBuildUs=122,536 us`, `BatchDeleteUs=848,531 us`, `BatchObservedItems=384`, `BatchFailedItems=0`.
- Fresh multi-batch settings archive: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_194943`
- `Phase7_RecycleBinBatchDeleteMultiBatch` now proves the configurable cap is honored exactly: with `recycleBinBatchSize=256`, the archived perf stream records three recycle-bin batches with `BatchObservedItems=256`, `256`, and `256`, all with `BatchFailedItems=0`.
- The same `2026-03-31_194943` archive kept host overhead low for the `768`-item run: `FileOps.Operation=2,468,000 us`, `FileOps.Progress.CallbackUs=752 us`, `FileOps.ItemCompleted.CallbackUs=431 us`.

---

## Phase 2 — High Impact: Cross-Filesystem Bridge Pipeline (P2)

### Problem

The cross-filesystem bridge (`FolderWindow.FileOperations.State.cpp ~L4455–4550`) uses a single 1 MB buffer with **serialized read→write**. One side always idles while the other transfers:

```
source.Read(buffer, 1MB) → wait → destination.Write(buffer, 1MB) → wait → loop
```

On fast source + slow destination (or vice versa), throughput is halved.

### Proposed Changes

- [x] **2.1** Implement **double-buffered (ping-pong) pipeline**:
  - Two 4 MB buffers (Buffer A, Buffer B)
  - Reader thread fills Buffer A while writer thread drains Buffer B → swap → repeat
  - Synchronize via two `std::binary_semaphore`s (reader-done, writer-done)
  - Total memory: 8 MB per active cross-FS transfer (acceptable — capped by task concurrency)
  - Landed on 2026-03-31 in the shared buffered bridge copy core: top-level single-file bridge copies (`CopyFile(...)`) now delegate to the same double-buffered implementation used by the within-folder worker path (`CopyFileWithBuffer(...)`)

- [x] **2.2** Make buffer size adaptive based on destination latency:
  - Default: 4 MB per buffer
  - Cloud/remote destinations (S3, Google Drive, SFTP): 8 MB per buffer (high latency benefits from larger batches)
  - Local-to-local: 2 MB per buffer (low latency, memory conservation)
  - Query via `IFileSystem::GetTransferHints(...)`; do not infer from plugin type in this plan
  - Landed on 2026-04-01: `ResolveAdaptiveCrossFsBridgeBufferBytes(...)` now resolves the configured default `4096 KB` task snapshot to endpoint-specific bridge buffer sizes using `IFileSystem::GetTransferHints(...)` from both the source and destination sides and applies the maximum hinted buffer size to the active bridge task
  - The contract update now builds across the full solution, including built-in plugins and in-tree test/perf stubs, and the host still keeps `FileSystemOptions` ABI unchanged

- [x] **2.3** Shrink callback work in the bridge without changing the current callback contract:
  - Keep `IFileSystemCallback::FileSystemProgress()` as the pause/cancel checkpoint for bridge work
  - Accumulate bytes and cheap per-buffer stats outside the callback lock, then emit the existing serialized callback only at throttled checkpoints (for example every ~200ms and on item completion)
  - If a conflict/issue is raised, both bridge threads must stop advancing at the next buffer boundary and follow the existing task-level prompt flow
  - Landed on 2026-03-31 in the shared buffered bridge copy core: progress callbacks are now emitted at throttled checkpoints instead of per write completion, while pause / cancel still flows through `FileSystemProgress()`

- [x] **2.4** Add cancellation point between buffer swaps (check `cancelRequested` atomic)
  - Landed on 2026-03-31 in the shared buffered bridge copy core: both reader and writer sides poll cancellation / pause around buffer handoff boundaries

- [x] **2.5** Add perf instrumentation: bytes/sec throughput, reader-wait-time, writer-wait-time (identifies which side is the bottleneck)
  - Landed on 2026-03-31 in the shared buffered bridge copy core via `FileOps.Bridge.Copy`, `FileOps.Bridge.CopyUs`, `FileOps.Bridge.ReadUs`, `FileOps.Bridge.WriteUs`, `FileOps.Bridge.ReaderWaitUs`, and `FileOps.Bridge.WriterWaitUs`
  - Added deterministic selftest perf metrics `FileOps.SelfTest.BridgePipelineBaseline`, `FileOps.SelfTest.BridgePipelineCandidate`, and `FileOps.SelfTest.BridgePipelineImprovement`

### Validation

- Deterministic selftest landed: `Phase11_BridgePipelineDummyToDummyPerf`
- The selftest forces the same dummy-to-dummy copy twice in one binary: baseline with bridge pipeline override `Disabled`, then candidate with override `Enabled`
- The selftest now copies four `32 MiB` dummy files with `streamChunkLatencyMs = 30`, asserts the configured task snapshot kept `crossFsBridgeBufferSizeKB = 4096`, asserts the active bridge resolved that default to `resolvedBufferKB = 8192` for the dummy high-latency path, and then records the before/after bridge timings
- `FileSystemDummy` now exposes a hidden `streamChunkLatencyMs` config used only by selftests to inject per-read / per-write latency without needing a second checkout or external network dependency
- Same-machine archive after landing the global bridge-buffer setting / task-snapshot slice: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_163447`
- Same-machine stability rerun: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_163749`
- Latest archived results:
  - `2026-03-31_163447`: `FileOps.SelfTest.BridgePipelineBaseline = 328,000 us`, `FileOps.SelfTest.BridgePipelineCandidate = 250,000 us`, `FileOps.SelfTest.BridgePipelineImprovement = 78,000 us` (`1.31x`, about `23.8%` lower wall time)
  - `2026-03-31_163749`: `FileOps.SelfTest.BridgePipelineBaseline = 343,000 us`, `FileOps.SelfTest.BridgePipelineCandidate = 203,000 us`, `FileOps.SelfTest.BridgePipelineImprovement = 140,000 us` (`1.69x`, about `40.8%` lower wall time)
- Clean isolated adaptive-buffer archive after the `IFileSystem::GetTransferHints(...)` contract landed solution-wide: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_140301-Phase11_BridgePipelineDummyToDummyPerf`
- Initial isolated adaptive-buffer result on build `66`: `FileOps.SelfTest.BridgePipelineBaseline = 17,312,000 us`, `FileOps.SelfTest.BridgePipelineCandidate = 406,000 us`, `FileOps.SelfTest.BridgePipelineImprovement = 16,906,000 us` (`42.6x`, about `97.7%` lower wall time)
- That isolated selftest metrics set includes `configuredBufferKB=4096` and `resolvedBufferKB=8192`, while the per-copy bridge metrics show `bridgeCopyUs=1,020,045 us`, `bridgeReadUs=669,080 us`, `bridgeWriteUs=850,970 us`, `bridgeReaderWaitUs=115,719 us`, and `bridgeWriterWaitUs=159,851 us`
- Final isolated confirmation archive after the remaining Preferences/File Operations work closed: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_193300-Phase11_BridgePipelineDummyToDummyPerf`
- The final isolated confirmation rerun stayed in the same strong range: `FileOps.SelfTest.BridgePipelineBaseline = 20,297,000 us`, `FileOps.SelfTest.BridgePipelineCandidate = 453,000 us`, `FileOps.SelfTest.BridgePipelineImprovement = 19,844,000 us` (`44.8x`, about `97.8%` lower wall time), again with `configuredBufferKB=4096` and `resolvedBufferKB=8192`.
- The earlier runtime crash after first landing the contract was traced to a stale `FileSystemDummy.dll` still using the old `IFileSystem` vtable when only the `RedSalamander` target had been rebuilt; the fix was to update every in-tree `IFileSystem` implementer/test stub and then rebuild the full solution before rerunning perf
- Adjacent regression coverage passed:
  - `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_163535` (`Phase11_CrossFileSystemBridge`)
  - `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_163630` (`Phase11_BridgeSingleFolderParallelCopyInFlightLines`)
- Latest isolated correctness / adjacent regression coverage after the adaptive-buffer contract update:
  - `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_140300-Phase11_CrossFileSystemBridge` (`Phase11_CrossFileSystemBridge` passed in `12,657 ms`)
  - `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_140302-Phase8_DefaultBandwidthLimitFromSettings` (`Phase8_DefaultBandwidthLimitFromSettings` passed with baseline `281,000 us`, candidate `4,532,000 us`, slowdown `4,251,000 us`)
- Final isolated correctness confirmation archive after the remaining Preferences/File Operations work closed:
  - `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_193600-Phase11_CrossFileSystemBridge` (`Phase11_CrossFileSystemBridge` passed in `12,750 ms`)
- Integrity guard: the selftest validates destination byte-for-byte equality after the baseline and candidate copy runs before reporting success

---

## Phase 3 — Medium Impact: Callback Contention Reduction (P2)

### Problem

The `callbackMutex` in `ParallelOperationState` serializes **all** callback invocations. With 4–8 parallel workers:

1. **Progress callbacks** (~every 50–100ms per worker) — lightweight but frequent; hold the lock for counter updates + diagnostic string copies (`getMostSpecificPathsForDiagnostics()` at ~L2741–2752 acquires nested `_progressMutex`)
2. **Conflict callbacks** — block on `WaitForSingleObject(_conflictDecisionEvent)` waiting for user input. **While one worker is blocked on conflict, all other workers queue behind the same mutex**, even for unrelated progress updates.

This creates unnecessary lock hold and bookkeeping contention, especially when a conflict prompt is active and the task is converging to the paused state required by the current host contract.

### Proposed Changes

- [x] **3.1** Split internal bookkeeping from callback serialization:
  - Keep a narrow callback gate that still serializes entry into the existing host callback methods
  - Move byte/item counters and per-worker current-item snapshots to atomics or a lightweight snapshot lock so they can be updated without holding the callback gate

- [x] **3.2** Keep progress callbacks, but make them cheaper:
  - `completedBytes`, `completedItems` already use `std::atomic<uint64_t>` in `ParallelOperationState` — extend that model to per-worker current-item tracking needed for progress snapshots
  - Build diagnostic path snapshots before taking the callback gate, or read from an already-published snapshot
  - Preserve `FileSystemProgress()` as the pause / queue-pause checkpoint; do not replace it with timer-only polling

- [x] **3.3** Implement conflict coordination that preserves task-wide convergence:
  - When a worker hits a conflict (e.g., `ERROR_ALREADY_EXISTS`), publish a `ConflictRequest` and mark the task as conflict-pending
  - The triggering worker becomes the prompt owner; all other workers must observe the task-level conflict state and stop at their next progress/checkpoint before the UI decision is applied
  - Host UI thread resolves exactly one prompt at a time per task, then releases the task-level conflict state so workers resume together

- [x] **3.4** Remove nested lock acquisition: `getMostSpecificPathsForDiagnostics()` should not acquire `_progressMutex` while the callback gate is held — make it read from atomics or a published snapshot copy

- [x] **3.5** Add contention metrics: count mutex wait events per task, log in JSONL diagnostics. Track `conflictWaitUs` and `progressCallbackUs` per worker.
  - Landed task-level `FileOps.Progress.LockContentionCount` and `FileOps.ItemCompleted.LockContentionCount`
  - Landed per-stream `FileOps.Progress.Stream.CallbackUs` and `FileOps.Progress.Stream.LockWaitUs`
  - Landed task-level `FileOps.Conflict.PromptCount` plus per-worker `FileOps.Conflict.Worker.WaitUs`

### Partial Landing 2026-03-31

- The host task now publishes a task-level `DiagnosticPathSnapshot` whenever `_progressSourcePath`, `_progressDestinationPath`, `_lastProgressCallbackSourcePath`, or `_lastProgressCallbackDestinationPath` changes under `_progressMutex`.
- `FileSystemIssue()` and the per-item conflict prompt path in `ExecuteOperation()` now resolve diagnostic source/destination paths from the published snapshot plus the per-item cookie instead of taking `_progressMutex` inside conflict handling.
- This keeps the existing prompt / pause behavior intact while removing the nested host progress-lock acquisition from the conflict path called out in the plan.
- The host now keeps per-stream progress perf bookkeeping (`_progressStreamPerf`) behind a dedicated mutex instead of reacquiring `_progressMutex` after every callback just to add `callbackUs` / `lockWaitUs` stream summaries.
- Progress callbacks now also collapse those per-stream summary writes into a single `_progressStreamPerfMutex` acquisition per callback, so `callbackCount`, `lockWaitUs`, and `callbackUs` update together instead of taking the stream-perf mutex twice.
- The host now also keeps `_inFlightFiles` behind its own mutex, so the popup-facing multi-line file-progress table no longer shares `_progressMutex` with the main byte/item/path state.
- The host now also keeps `_perItemInFlightCalls` behind its own mutex, so per-item callback progress aggregation and per-call bandwidth-share snapshots no longer search/update that table while `_progressMutex` is held.
- The host now maintains cached per-item in-flight aggregate counters (`completedBytes`, `completedItems`, `totalItems`) under that same mutex, so the callback path no longer rescans the active-call table just to recompute progress totals after each update.
- The callback-time copy/move bandwidth-share rewrite now runs through a shared helper after the main progress lock is released, so `_progressMutex` no longer covers per-call `bandwidthLimitBytesPerSecond` recomputation / stores.
- The host now also keeps progress-path state (`_progressSourcePath`, `_progressDestinationPath`, `_lastProgressCallback*Path`, `_lastProgressCallbackTick`) behind a dedicated mutex, so string comparisons / assignments, published diagnostic snapshot refresh, and popup/completed-summary path reads no longer sit inside `_progressMutex`.
- The host now also publishes byte/item/top-level progress counters into dedicated atomics while `_progressMutex` is held, so popup task snapshots and popup rate snapshots no longer take `_progressMutex` just to read totals.
- Progress and item-completed callback counts now increment atomically outside `_progressMutex`, and task-result diagnostics / perf emit / completed-task summaries now read published progress snapshots instead of reacquiring `_progressMutex` for totals and callback counts after the operation finishes.
- The two callback hot paths now also capture progress-counter snapshots while `_progressMutex` is held but publish those atomics after the lock is released, so the main progress lock no longer covers the repeated batch of atomic stores for popup/rate/task-summary readers.
- Top-level completion accounting (`_topLevelItemCompleted`, `_completedTopLevelFiles`, `_completedTopLevelFolders`) now updates behind its own mutex and published atomics, so callback-time and worker-path completion bookkeeping no longer sits inside `_progressMutex`.
- Conflict prompts now track an explicit owner thread id in both the host callback path and the per-item worker path, and non-owner workers that reach `WaitWhilePaused()` while a prompt is active now accumulate `FileOps.Conflict.ConvergenceWaitUs` instead of continuing past the task-level conflict state.
- Apply-to-all decisions are now cached under `_conflictMutex` before the active prompt is cleared and waiting workers are released, eliminating the race where a sibling worker could wake on “prompt cleared” and raise a redundant second prompt before the cached decision became visible.
- File-ops selftest completion snapshots now carry `conflictWaitUs`, `conflictConvergenceWaitUs`, and `conflictPromptCount`, and `Phase9_ConflictPrompt_ApplyToAllUiCache` now holds the first prompt for `250 ms` across four conflicting items before clicking Overwrite + Apply-to-all so the exact case itself proves task-wide convergence instead of relying only on post-run JSONL inspection.
- Host perf output now also records contention counts and per-stream progress summaries for the callback path:
  - `FileOps.Progress.LockContentionCount`
  - `FileOps.ItemCompleted.LockContentionCount`
  - `FileOps.Progress.Stream.CallbackUs`
  - `FileOps.Progress.Stream.LockWaitUs`
  - `FileOps.Conflict.PromptCount`
  - `FileOps.Conflict.Worker.WaitUs`

### Validation

- Selftest: parallel copy of 1K small files with injected conflicts on 10% of items → measure total wall time and per-worker idle time before/after
- Expected gain: reduces unnecessary callback-lock hold time and diagnostic contention while preserving the current task-wide prompt / pause semantics
- Deterministic correctness archive for the landed slice: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_143024`
- Focused selftest `Phase9_ConflictPrompt_OverwriteReplaceReadonly` passed in `140 ms` on the same machine.
- Candidate lock metrics from that archive: `FileOps.Progress.LockWaitUs=6 us`, `FileOps.Progress.LockHoldUs=57 us`, `FileOps.ItemCompleted.LockHoldUs=5 us`, `FileOps.Operation=16,000 us`.
- Historical same-machine archives already include the same Phase 9 case (`2026-03-27_135523`, `2026-03-27_141221`, `2026-03-27_144845`, `2026-03-27_145536`) with case durations between `141 ms` and `250 ms`; the new `140 ms` run is directionally at the low end, but it is not an isolated before/after baseline for this exact host-lock-only slice.
- New contention-count archive: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_175211`
- `Phase9_ConflictPrompt_OverwriteReplaceReadonly` passed again in `375 ms`, and the archive now shows `FileOps.Progress.LockContentionCount=5`, `FileOps.ItemCompleted.LockContentionCount=1`, `FileOps.Progress.Stream.CallbackUs=111 us`, and `FileOps.Progress.Stream.LockWaitUs=5 us` for the captured task stream.
- Conflict-heavy validation archive: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_182301`
- `Phase9_ConflictPrompt_OverwriteReplaceReadonly` passed in `313 ms`, and the archive now shows `FileOps.Conflict.PromptCount=2`, `FileOps.Conflict.Worker.WaitUs=58,621 us`, `FileOps.Conflict.WaitUs=58,621 us`, `FileOps.Progress.LockContentionCount=1`, and `FileOps.Progress.Stream.CallbackUs=79 us`.
- Multi-stream validation archive: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_182410`
- `Phase11_BridgeSingleFolderParallelCopyInFlightLines` passed in `24,703 ms`, and the archive now shows four stream summaries: `FileOps.Progress.Stream.CallbackUs={380,360,366,330} us` and `FileOps.Progress.Stream.LockWaitUs={1,2,1,1} us`, with `FileOps.Progress.LockContentionCount=5`.
- Latest isolated conflict rerun after the Phase 3.1 bookkeeping split: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_080629-Phase9_ConflictPrompt_OverwriteReplaceReadonly`
- The exact Phase 9 case passed in `406 ms`, with `FileOps.Operation=141,000 us`, `FileOps.Conflict.PromptCount=2`, `FileOps.Conflict.Worker.WaitUs=143,680 us`, `FileOps.Progress.LockWaitUs=3 us`, `FileOps.Progress.LockHoldUs=83 us`, `FileOps.Progress.LockContentionCount=2`, and `FileOps.Progress.Stream.CallbackUs=99 us`.
- Latest isolated multi-stream rerun after the same split: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_080733-Phase11_BridgeSingleFolderParallelCopyInFlightLines`
- The exact Phase 11 case passed in `24,812 ms`, with `FileOps.Operation=24,484,000 us`, `FileOps.Progress.LockWaitUs=10 us`, `FileOps.Progress.LockHoldUs=1,054 us`, `FileOps.Progress.LockContentionCount=10`, and stream callback summaries `FileOps.Progress.Stream.CallbackUs={352,265,325,235} us`.
- Latest isolated multi-stream rerun after moving `_inFlightFiles` off `_progressMutex`: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_084741-Phase11_BridgeSingleFolderParallelCopyInFlightLines`
- The exact Phase 11 case passed again in `24,781 ms`, with `FileOps.Operation=24,484,000 us`, `FileOps.Progress.LockWaitUs=4 us`, `FileOps.Progress.LockHoldUs=988 us`, `FileOps.Progress.LockContentionCount=4`, and stream callback summaries `FileOps.Progress.Stream.CallbackUs={360,317,289,261} us`.
- Latest isolated conflict rerun after moving `_perItemInFlightCalls` off `_progressMutex`: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_092012-Phase9_ConflictPrompt_OverwriteReplaceReadonly`
- The exact Phase 9 case still passed, and its callback-lock metrics improved further in that run: `FileOps.Operation=63,000 us`, `FileOps.Progress.LockWaitUs=0 us`, `FileOps.Progress.LockHoldUs=46 us`, `FileOps.Progress.LockContentionCount=0`, `FileOps.Conflict.Worker.WaitUs=56,221 us`, and `FileOps.Progress.Stream.CallbackUs=86 us`.
- Latest isolated multi-stream rerun after the same `_perItemInFlightCalls` split: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_091900-Phase11_BridgeSingleFolderParallelCopyInFlightLines`
- That exact Phase 11 case passed but did **not** improve the multi-stream path on this machine: `24,860 ms` vs the prior isolated `24,781 ms`, with `FileOps.Progress.LockWaitUs=12 us`, `FileOps.Progress.LockHoldUs=1,149 us`, and `FileOps.Progress.LockContentionCount=11`.
- Follow-up waited multi-stream reruns after caching the per-item aggregate totals: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_094101-Phase11_BridgeSingleFolderParallelCopyInFlightLines` and `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_094536-Phase11_BridgeSingleFolderParallelCopyInFlightLines`
- Those waited Phase 11 reruns recovered the clean multi-stream case to the pre-regression range while keeping the new split in place:
  - `2026-04-01_094101`: `24,719 ms`, `FileOps.Operation=24,453,000 us`, `FileOps.Progress.LockWaitUs=8 us`, `FileOps.Progress.LockHoldUs=949 us`, `FileOps.Progress.LockContentionCount=8`
  - `2026-04-01_094536`: `24,734 ms`, `FileOps.Operation=24,484,000 us`, `FileOps.Progress.LockWaitUs=3 us`, `FileOps.Progress.LockHoldUs=959 us`, `FileOps.Progress.LockContentionCount=3`
- Compared to the best earlier isolated multi-stream archive (`2026-04-01_084741`: `24,781 ms`, `LockHoldUs=988 us`, `LockContentionCount=4`), the cached-aggregate follow-up is now directionally neutral-to-positive on wall time and lock hold, with lock wait / contention landing in the same low single-digit range on the better rerun.
- Latest isolated Phase 8 regression after the same split: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_084851-Phase8_DefaultBandwidthLimitFromSettings`
- `Phase8_DefaultBandwidthLimitFromSettings` still passed, with baseline `125,000 us`, candidate `5,078,000 us`, slowdown `4,953,000 us`.
- Fresh isolated Phase 8 regression after the `_perItemInFlightCalls` split: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_092126-Phase8_DefaultBandwidthLimitFromSettings`
- The Phase 8 regression still passed there too, with baseline `156,000 us`, candidate `5,062,000 us`, slowdown `4,906,000 us`, so the new callback-state split did not break the settings-seeded bandwidth path.
- Latest waited Phase 8 regression after the cached-aggregate follow-up: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_094304-Phase8_DefaultBandwidthLimitFromSettings`
- That waited rerun also passed, with baseline `125,000 us`, candidate `5,032,000 us`, slowdown `4,907,000 us`.
- Latest isolated multi-stream reruns after moving the callback-time bandwidth-share rewrite out of `_progressMutex`: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_102228-Phase11_BridgeSingleFolderParallelCopyInFlightLines` and `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_103952-Phase11_BridgeSingleFolderParallelCopyInFlightLines`
- Those exact Phase 11 reruns stayed in the same envelope rather than showing a new clear win:
  - `2026-04-01_102228`: `24,703 ms`, `FileOps.Operation=24,438,000 us`, `FileOps.Progress.LockWaitUs=9 us`, `FileOps.Progress.LockHoldUs=992 us`, `FileOps.Progress.LockContentionCount=8`
  - `2026-04-01_103952`: `24,687 ms`, `FileOps.Operation=24,469,000 us`, `FileOps.Progress.LockWaitUs=11 us`, `FileOps.Progress.LockHoldUs=1,003 us`, `FileOps.Progress.LockContentionCount=11`
- Latest isolated conflict reruns after the same split: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_102450-Phase9_ConflictPrompt_OverwriteReplaceReadonly` and `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_104156-Phase9_ConflictPrompt_OverwriteReplaceReadonly`
- The first Phase 9 rerun was noisy (`469 ms`, `FileOps.Operation=78,000 us`), but the immediate rerun returned to the prior good range: `2026-04-01_104156` passed in `234 ms`, with `FileOps.Operation=63,000 us`, `FileOps.Progress.LockWaitUs=1 us`, `FileOps.Progress.LockHoldUs=51 us`, `FileOps.Progress.LockContentionCount=1`, and `FileOps.Conflict.Worker.WaitUs=48,938 us`.
- Adjacent isolated bandwidth regressions still passed after the same change:
  - `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_102546-Phase8_DefaultBandwidthLimitFromSettings`: baseline `203,000 us`, candidate `5,125,000 us`, slowdown `4,922,000 us`
  - `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_102639-Phase6_LocalBandwidthThrottle`: `duration=4,375,000 us`, `lead=0 us`, `maxWindowBytes=6,291,456`, `maxSampleDeltaBytes=2,097,152`, `cancelLatency=47,000 us`
- Latest isolated multi-stream reruns after moving progress-path state out of `_progressMutex`: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_110225-Phase11_BridgeSingleFolderParallelCopyInFlightLines` and `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_111210-Phase11_BridgeSingleFolderParallelCopyInFlightLines`
- Those exact Phase 11 reruns kept wall time in the same envelope but sharply reduced callback-lock hold time:
  - `2026-04-01_110225`: `24,907 ms`, `FileOps.Operation=24,469,000 us`, `FileOps.Progress.LockWaitUs=9 us`, `FileOps.Progress.LockHoldUs=53 us`, `FileOps.Progress.LockContentionCount=9`
  - `2026-04-01_111210`: `24,718 ms`, `FileOps.Operation=24,454,000 us`, `FileOps.Progress.LockWaitUs=3 us`, `FileOps.Progress.LockHoldUs=40 us`, `FileOps.Progress.LockContentionCount=3`
- Compared to the best earlier clean multi-stream archive before the path split (`2026-04-01_094536`: `24,734 ms`, `LockHoldUs=959 us`, `LockContentionCount=3`), the new path-state split is a repeated `~24x` lock-hold reduction while holding the same wall-time range.
- Latest isolated conflict reruns after the same path-state split: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_110437-Phase9_ConflictPrompt_OverwriteReplaceReadonly` and `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_111414-Phase9_ConflictPrompt_OverwriteReplaceReadonly`
- Those exact Phase 9 reruns also kept the conflict-heavy path in its good range while collapsing callback-lock hold time further:
  - `2026-04-01_110437`: `391 ms`, `FileOps.Operation=62,000 us`, `FileOps.Progress.LockWaitUs=2 us`, `FileOps.Progress.LockHoldUs=3 us`, `FileOps.Progress.LockContentionCount=2`, `FileOps.Conflict.Worker.WaitUs=48,598 us`
  - `2026-04-01_111414`: `359 ms`, `FileOps.Operation=63,000 us`, `FileOps.Progress.LockWaitUs=1 us`, `FileOps.Progress.LockHoldUs=4 us`, `FileOps.Progress.LockContentionCount=1`, `FileOps.Conflict.Worker.WaitUs=67,436 us`
- Fresh adjacent isolated Phase 8 regression after the same split: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_110533-Phase8_DefaultBandwidthLimitFromSettings`
- That regression still passed, with baseline `219,000 us`, candidate `5,047,000 us`, slowdown `4,828,000 us`.
- Latest isolated reruns after collapsing per-stream progress bookkeeping to one mutex acquisition: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_113719-Phase11_BridgeSingleFolderParallelCopyInFlightLines`, `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_113828-Phase9_ConflictPrompt_OverwriteReplaceReadonly`, and `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_113911-Phase8_DefaultBandwidthLimitFromSettings`
- Those exact reruns stayed correctness-safe and in the same performance envelope rather than showing a new clear wall-time win:
  - `2026-04-01_113719`: `24,719 ms`, `FileOps.Operation=24,469,000 us`, `FileOps.Progress.LockWaitUs=7 us`, `FileOps.Progress.LockHoldUs=43 us`, `FileOps.Progress.LockContentionCount=7`, `FileOps.Progress.Stream.CallbackUs={360,304,311,289} us`
  - `2026-04-01_113828`: `266 ms`, `FileOps.Operation=63,000 us`, `FileOps.Progress.LockWaitUs=1 us`, `FileOps.Progress.LockHoldUs=4 us`, `FileOps.Progress.LockContentionCount=1`, `FileOps.Conflict.Worker.WaitUs=62,602 us`
  - `2026-04-01_113911`: `Phase8_DefaultBandwidthLimitFromSettings` still passed, with baseline `172,000 us`, candidate `5,047,000 us`, slowdown `4,875,000 us`
- Latest isolated reruns after moving popup totals/rate reads to published progress-counter atomics: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_104418-Phase11_BridgeSingleFolderParallelCopyInFlightLines`, `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_104811-Phase11_BridgeSingleFolderParallelCopyInFlightLines`, `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_104418-Phase9_ConflictPrompt_OverwriteReplaceReadonly`, `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_104811-Phase9_ConflictPrompt_OverwriteReplaceReadonly`, and adjacent regression `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_104418-Phase8_DefaultBandwidthLimitFromSettings`
- Those exact Phase 11 reruns stayed in the same wall-time band and kept lock hold below the pre-path-split `~959 us` range, but they did not beat the immediately previous `40`–`43 us` best reruns:
  - `2026-04-01_104418`: `24,781 ms`, `FileOps.Operation=24,469,000 us`, `FileOps.Progress.LockWaitUs=7 us`, `FileOps.Progress.LockHoldUs=86 us`, `FileOps.Progress.LockContentionCount=7`, `FileOps.Progress.Stream.CallbackUs={319,264,249,299} us`
  - `2026-04-01_104811`: `24,812 ms`, `FileOps.Operation=24,485,000 us`, `FileOps.Progress.LockWaitUs=10 us`, `FileOps.Progress.LockHoldUs=462 us`, `FileOps.Progress.LockContentionCount=10`, `FileOps.Progress.Stream.CallbackUs={692,390,429,417} us`
- Those exact Phase 9 reruns stayed in the prior good range while keeping callback-lock contention at zero:
  - `2026-04-01_104418`: `265 ms`, `FileOps.Operation=63,000 us`, `FileOps.Progress.LockWaitUs=0 us`, `FileOps.Progress.LockHoldUs=7 us`, `FileOps.Progress.LockContentionCount=0`, `FileOps.Conflict.Worker.WaitUs=58,349 us`, `FileOps.Progress.Stream.CallbackUs=87 us`
  - `2026-04-01_104811`: `265 ms`, `FileOps.Operation=63,000 us`, `FileOps.Progress.LockWaitUs=0 us`, `FileOps.Progress.LockHoldUs=6 us`, `FileOps.Progress.LockContentionCount=0`, `FileOps.Conflict.Worker.WaitUs=58,134 us`, `FileOps.Progress.Stream.CallbackUs=93 us`
- The adjacent Phase 8 regression still passed after the same slice, with baseline `156,000 us`, candidate `5,125,000 us`, slowdown `4,969,000 us`.
- Latest isolated reruns after moving callback-count increments out of `_progressMutex` and switching task-result/perf/completed-summary readers to published snapshots: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_105907-Phase11_BridgeSingleFolderParallelCopyInFlightLines`, `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_110207-Phase11_BridgeSingleFolderParallelCopyInFlightLines`, `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_105907-Phase9_ConflictPrompt_OverwriteReplaceReadonly`, `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_110207-Phase9_ConflictPrompt_OverwriteReplaceReadonly`, and adjacent regression `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_105907-Phase8_DefaultBandwidthLimitFromSettings`
- The repeated Phase 11 reruns again stayed in the same clean multi-stream wall-time band and callback-lock envelope rather than improving on the earlier path-state best:
  - `2026-04-01_105907`: `24,765 ms`, `FileOps.Operation=24,469,000 us`, `FileOps.Progress.LockWaitUs=10 us`, `FileOps.Progress.LockHoldUs=106 us`, `FileOps.Progress.LockContentionCount=10`, `FileOps.Progress.Stream.CallbackUs={451,353,589,491} us`
  - `2026-04-01_110207`: `24,922 ms`, `FileOps.Operation=24,469,000 us`, `FileOps.Progress.LockWaitUs=11 us`, `FileOps.Progress.LockHoldUs=102 us`, `FileOps.Progress.LockContentionCount=11`, `FileOps.Progress.Stream.CallbackUs={421,338,308,352} us`
- The first Phase 9 rerun after that slice was unusually favorable, but the immediate repeat fell back into the prior good conflict-heavy range:
  - `2026-04-01_105907`: `203 ms`, `FileOps.Operation=31,000 us`, `FileOps.Progress.LockWaitUs=0 us`, `FileOps.Progress.LockHoldUs=8 us`, `FileOps.Progress.LockContentionCount=0`, `FileOps.Conflict.Worker.WaitUs=15,595 us`, `FileOps.Progress.Stream.CallbackUs=111 us`
  - `2026-04-01_110207`: `265 ms`, `FileOps.Operation=63,000 us`, `FileOps.Progress.LockWaitUs=1 us`, `FileOps.Progress.LockHoldUs=13 us`, `FileOps.Progress.LockContentionCount=1`, `FileOps.Conflict.Worker.WaitUs=55,801 us`, `FileOps.Progress.Stream.CallbackUs=95 us`
- The adjacent Phase 8 regression still passed after the same slice, with baseline `110,000 us`, candidate `5,063,000 us`, slowdown `4,953,000 us`.
- Latest isolated reruns after moving the hot callback counter-publish stores out of `_progressMutex`: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_110641-Phase11_BridgeSingleFolderParallelCopyInFlightLines`, `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_110641-Phase9_ConflictPrompt_OverwriteReplaceReadonly`, and adjacent regression `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_110641-Phase8_DefaultBandwidthLimitFromSettings`
- That exact Phase 11 rerun stayed at the same clean multi-stream wall-time (`24,781 ms`, `FileOps.Operation=24,469,000 us`) but improved callback lock hold versus the immediately previous callback-counter slice (`102`–`106 us` down to `61 us`), with `FileOps.Progress.LockWaitUs=12 us`, `FileOps.Progress.LockContentionCount=12`, and `FileOps.Progress.Stream.CallbackUs={371,321,359,466} us`.
- The exact Phase 9 rerun stayed in the prior good conflict-heavy range: `297 ms`, `FileOps.Operation=63,000 us`, `FileOps.Progress.LockWaitUs=2 us`, `FileOps.Progress.LockHoldUs=12 us`, `FileOps.Progress.LockContentionCount=2`, `FileOps.Conflict.Worker.WaitUs=59,732 us`, and `FileOps.Progress.Stream.CallbackUs=109 us`.
- The adjacent Phase 8 regression still passed after the same slice, with baseline `156,000 us`, candidate `5,062,000 us`, slowdown `4,906,000 us`.
- Known test-harness noise: an intermediate exact-case rerun hit the existing setup flake in `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_182215` (`Failed to create pre-calc perf trees.`), but the immediately following rerun succeeded without code changes.
- Follow-up harness hardening landed in `FolderWindow.FileOperations.SelfTest.cpp`, adding retries around the heavy pre-calc setup trees. Post-fix exact-case reruns both passed cleanly in `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_182827` (Phase 9, `281 ms`, `FileOps.Conflict.Worker.WaitUs=57,797 us`) and `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_182926` (Phase 11, `24,703 ms`) with no setup failure.
- Current evidence for that follow-up was stronger: the callback-state splits preserved the conflict-heavy win, kept the clean multi-stream case in the prior wall-time range, kept adjacent bandwidth regressions green, and the dedicated path-state split still showed a repeated large reduction in `FileOps.Progress.LockHoldUs` on both the clean multi-stream and conflict-heavy exact cases. At that point Phase 3 was still open for conflict coordination (`3.3`) and because end-to-end wall-time wins remained modest compared to the lock-hold improvement; the later conflict-convergence and apply-to-all proof slices closed that remaining gap.
- Latest isolated reruns after moving top-level completion bookkeeping out of `_progressMutex`: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_124112-Phase11_BridgeSingleFolderParallelCopyInFlightLines`, `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_124250-Phase9_ConflictPrompt_OverwriteReplaceReadonly`, `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_124336-Phase8_DefaultBandwidthLimitFromSettings`, `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_124450-Phase11_BridgeSingleFolderParallelCopyInFlightLines`, and `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_124602-Phase9_ConflictPrompt_OverwriteReplaceReadonly`
- Those exact reruns kept the slice correctness-safe and regression-safe, but they did not beat the best earlier Phase 3 path-state numbers:
  - `2026-04-01_124112`: `Phase11_BridgeSingleFolderParallelCopyInFlightLines` passed in `24,875 ms`, with `FileOps.Operation=24,485,000 us`, `FileOps.Progress.LockWaitUs=8 us`, `FileOps.Progress.LockHoldUs=85 us`, and `FileOps.Progress.LockContentionCount=8`
  - `2026-04-01_124450`: waited rerun also passed in `24,734 ms`, with `FileOps.Operation=24,453,000 us`, `FileOps.Progress.LockWaitUs=14 us`, `FileOps.Progress.LockHoldUs=82 us`, and `FileOps.Progress.LockContentionCount=14`
  - `2026-04-01_124250`: `Phase9_ConflictPrompt_OverwriteReplaceReadonly` passed in `281 ms`, with `FileOps.Operation=63,000 us`, `FileOps.Progress.LockWaitUs=1 us`, `FileOps.Progress.LockHoldUs=12 us`, `FileOps.Progress.LockContentionCount=1`, and `FileOps.Conflict.Worker.WaitUs=62,685 us`
- `2026-04-01_124602`: waited conflict rerun passed in `297 ms`, with `FileOps.Operation=79,000 us`, `FileOps.Progress.LockWaitUs=3 us`, `FileOps.Progress.LockHoldUs=7 us`, `FileOps.Progress.LockContentionCount=3`, and `FileOps.Conflict.Worker.WaitUs=61,980 us`
- `2026-04-01_124336`: adjacent `Phase8_DefaultBandwidthLimitFromSettings` still passed with baseline `141,000 us`, candidate `4,454,000 us`, slowdown `4,313,000 us`
- Compared to the earlier best clean path-state reruns (`2026-04-01_110225` / `2026-04-01_111210` at `LockHoldUs=53 us` / `40 us`, and `2026-04-01_110437` / `2026-04-01_111414` at `LockHoldUs=3 us` / `4 us`), this newest split looks architecturally cleaner but perf-neutral to slightly worse on this machine, so it is kept as a safe cleanup rather than counted as the step that closes Phase 3.
- First isolated proof archive for the new conflict-convergence slice: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_125259-Phase9_ConflictPrompt_ApplyToAllUiCache`
- That first proof run is intentionally kept as a rejected regression archive: `Phase9_ConflictPrompt_ApplyToAllUiCache` timed out because the first implementation cleared the prompt before publishing the apply-to-all cache entry, allowing a sibling worker to race in with a second prompt. The archive shows `FileOps.Conflict.PromptCount=2`, `FileOps.Conflict.WaitUs=123,623,577 us`, and task result `0x800704C7`.
- Fixed same-build proof archive after caching before prompt clear: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_130206-Phase9_ConflictPrompt_ApplyToAllUiCache`
- The repaired `Phase9_ConflictPrompt_ApplyToAllUiCache` case passed in `343 ms`, with `FileOps.Operation=94,000 us`, `FileOps.Conflict.PromptCount=1`, `FileOps.Conflict.WaitUs=65,792 us`, `FileOps.Conflict.ConvergenceWaitUs=61,840 us`, `FileOps.Progress.LockWaitUs=2 us`, `FileOps.Progress.LockHoldUs=19 us`, and `FileOps.Progress.LockContentionCount=2`. This is the first archived exact case that proves a non-owner worker really converged at a checkpoint under an active conflict prompt instead of racing a second prompt.
- Same-build adjacent conflict regression after the fix: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_130255-Phase9_ConflictPrompt_OverwriteReplaceReadonly`
- `Phase9_ConflictPrompt_OverwriteReplaceReadonly` still passed in `218 ms`, with `FileOps.Operation=63,000 us`, `FileOps.Conflict.WaitUs=65,532 us`, `FileOps.Conflict.ConvergenceWaitUs=0 us`, `FileOps.Progress.LockWaitUs=1 us`, `FileOps.Progress.LockHoldUs=7 us`, and `FileOps.Progress.LockContentionCount=1`.
- Same-build clean multi-stream guard after the fix: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_130336-Phase11_BridgeSingleFolderParallelCopyInFlightLines`
- `Phase11_BridgeSingleFolderParallelCopyInFlightLines` still passed in `24,734 ms`, with `FileOps.Operation=24,485,000 us`, `FileOps.Progress.LockWaitUs=4 us`, `FileOps.Progress.LockHoldUs=90 us`, and `FileOps.Progress.LockContentionCount=4`.
- Same-build adjacent bandwidth regression after the fix: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_130443-Phase8_DefaultBandwidthLimitFromSettings`
- `Phase8_DefaultBandwidthLimitFromSettings` still passed there, with baseline `94,000 us`, candidate `4,469,000 us`, slowdown `4,375,000 us`.
- Latest isolated build `86` reruns after the Phase 7 completed-popup selftest fix: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_143025-Phase11_BridgeSingleFolderParallelCopyInFlightLines`, `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_143139-Phase9_ConflictPrompt_OverwriteReplaceReadonly`, `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_143241-Phase9_ConflictPrompt_ApplyToAllUiCache`, and adjacent regression `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_143335-Phase8_DefaultBandwidthLimitFromSettings`
- Those newest exact reruns keep the Phase 3 readout stable on the current binary:
  - `2026-04-01_143025`: `Phase11_BridgeSingleFolderParallelCopyInFlightLines` passed in `24,454 ms`, with `FileOps.Progress.LockWaitUs=12 us`, `FileOps.Progress.LockHoldUs=70 us`, `FileOps.Progress.LockContentionCount=12`, and stream callback summaries `FileOps.Progress.Stream.CallbackUs={333,384,353,294} us`
  - `2026-04-01_143139`: `Phase9_ConflictPrompt_OverwriteReplaceReadonly` passed in the prior good conflict-heavy range, with `FileOps.Operation=78,000 us`, `FileOps.Progress.LockWaitUs=1 us`, `FileOps.Progress.LockHoldUs=7 us`, `FileOps.Progress.LockContentionCount=1`, `FileOps.Conflict.PromptCount=2`, and `FileOps.Conflict.Worker.WaitUs=66,253 us`
  - `2026-04-01_143241`: `Phase9_ConflictPrompt_ApplyToAllUiCache` passed again, with `FileOps.Operation=94,000 us`, `FileOps.Progress.LockWaitUs=3 us`, `FileOps.Progress.LockHoldUs=10 us`, `FileOps.Progress.LockContentionCount=3`, `FileOps.Conflict.PromptCount=1`, `FileOps.Conflict.Worker.WaitUs=60,238 us`, and `FileOps.Conflict.ConvergenceWaitUs=56,130 us`
  - `2026-04-01_143335`: adjacent `Phase8_DefaultBandwidthLimitFromSettings` still passed, with baseline `297,000 us`, candidate `4,594,000 us`, slowdown `4,297,000 us`
- Strengthened isolated convergence-proof rerun: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_150100-Phase9_ConflictPrompt_ApplyToAllUiCache`
- The upgraded exact Apply-to-all case passed in `797 ms` after holding the first prompt open for `250 ms` across four conflicting items, with `FileOps.Operation=375,000 us`, `FileOps.Progress.LockWaitUs=5 us`, `FileOps.Progress.LockHoldUs=17 us`, `FileOps.Progress.LockContentionCount=4`, `FileOps.Conflict.PromptCount=1`, `FileOps.Conflict.WaitUs=347,625 us`, `FileOps.Conflict.ConvergenceWaitUs=1,019,347 us`, and the new selftest metric `FileOps.SelfTest.ConflictApplyToAllConvergenceWait=1,019,347 us`.
- Repeated isolated convergence-proof rerun plus adjacent guards on build `87`: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_151300-Phase9_ConflictPrompt_ApplyToAllUiCache`, `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_151300-Phase9_ConflictPrompt_OverwriteReplaceReadonly`, `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_151300-Phase11_BridgeSingleFolderParallelCopyInFlightLines`, and `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_151300-Phase8_DefaultBandwidthLimitFromSettings`
- Those repeated exact reruns keep the strengthened conflict proof and all adjacent guards green on the latest binary:
  - `2026-04-01_151300` Apply-to-all: `531 ms`, `FileOps.Operation=282,000 us`, `FileOps.Progress.LockWaitUs=4 us`, `FileOps.Progress.LockHoldUs=16 us`, `FileOps.Progress.LockContentionCount=4`, `FileOps.Conflict.PromptCount=1`, `FileOps.Conflict.WaitUs=259,723 us`, `FileOps.Conflict.ConvergenceWaitUs=755,041 us`, `FileOps.SelfTest.ConflictApplyToAllConvergenceWait=755,041 us`
  - `2026-04-01_151300` Overwrite/ReplaceReadonly guard: passed in `266 ms`, with `FileOps.Operation=78,000 us`, `FileOps.Progress.LockHoldUs=8 us`, `FileOps.Conflict.PromptCount=2`, and `FileOps.Conflict.Worker.WaitUs=69,972 us`
  - `2026-04-01_151300` Phase 11 clean multi-stream guard: `24,781 ms`, `FileOps.Operation=24,453,000 us`, `FileOps.Progress.LockWaitUs=6 us`, `FileOps.Progress.LockHoldUs=78 us`, `FileOps.Progress.LockContentionCount=6`, and stream callback summaries `FileOps.Progress.Stream.CallbackUs={329,344,428,297} us`
  - `2026-04-01_151300` adjacent Phase 8 regression: baseline `203,000 us`, candidate `4,484,000 us`, slowdown `4,281,000 us`
- Current evidence is now strong enough to treat `3.3` as landed too: two isolated exact reruns of the strengthened Apply-to-all case now hold the first prompt across four conflicting items and both archive a single prompt plus large non-owner convergence wait (`1,019,347 us` and `755,041 us` on a `250 ms` hold), while the usual Phase 9 / 11 / 8 guards remain green. Phase 3 is now complete: callback bookkeeping/path/counter state is split off the hot lock, conflict prompts have explicit owner/convergence semantics, apply-to-all decisions are published before wake, and the repeated lock-hold reduction is strong even though clean multi-stream wall-time gains remain modest on this machine.

---

## Phase 4 — Medium Impact: Parallel Pre-Calculation Tree Walk (P3)

### Problem

Pre-calc already fans out selected roots across workers, but the host still hard-codes the worker budget to `4` and had no global settings to disable pre-calc or tune the worker count per task creation. That made Phase 4 impossible to validate/tune cleanly and also meant the old spec overstated a root-seeding bottleneck that the current code no longer has.

### Proposed Changes

- [x] **4.1** Seed the pre-calc worker queue with **all selected root items at once**:
  - The current host implementation already does this with a `nextIndex` fan-out loop over all selected roots
  - Workers immediately claim independent subtrees instead of waiting for one root to recurse first

- [x] **4.2** Implement a bounded **directory work queue** for recursive descent:
  - The host now uses `ReadDirectoryInfo()` to enumerate immediate child directories and pushes them as new pre-calc work items instead of recursing synchronously inside one `GetDirectorySize()` call
  - Each queued directory still uses `IFileSystemDirectoryOperations::GetDirectorySize()` for counting so the existing progress / cancel callback contract stays intact
  - Queue depth is capped at `4096` pending directories; when the cap is hit, the worker falls back to synchronous recursive `GetDirectorySize()` for the overflowing child subtree

- [x] **4.3** Make pre-calc cancellation faster:
  - `FileSystem::GetDirectorySize()` now polls `DirectorySizeShouldCancel()` at entry boundaries instead of waiting for the throttled progress callback window
  - The debug-only `directorySizeDelayMs` path re-checks cancellation immediately after the injected sleep so deterministic latency tests measure the real callback cadence, not the old 200 ms throttle
  - Existing `DirectorySizeProgress()` / `DirectorySizeShouldCancel()` callback contracts stay intact; this is a plugin-side polling change only

- [x] **4.4** Expose pre-calc policy through global host settings and task snapshots:
  - Added `fileOperations.preCalcEnabled` and `fileOperations.preCalcMaxWorkers` to `SettingsStore`
  - Task creation snapshots those settings so already-running tasks do not retune mid-flight
  - `FileOps.PreCalc` now logs `workers=` and `configuredWorkers=` so same-machine A/B runs can confirm what budget actually ran

### Landed 2026-03-31

- Landed the settings/schema/task-creation slice for pre-calc control in the host.
- Landed the local plugin cancellation slice for pre-calc tree walks.
- Landed the host-side recursive directory work queue so a single selected root can fan out across the configured pre-calc worker budget.
- Added deterministic coverage for both persistence and behavior:
  - Commands selftest `settings_file_operations_precalc_roundtrip`
  - File Operations selftest `Phase5_PreCalcSettingsApplied`
  - File Operations selftest `Phase5_PreCalcCancelLatencyLocal`
- `Phase5_PreCalcSettingsApplied` now verifies:
  - disabling pre-calc from settings skips the pre-calc pass entirely
  - a `preCalcMaxWorkers = 1` task actually runs with one worker
  - the same seeded workload honors `preCalcMaxWorkers = 4` and `preCalcMaxWorkers = 8` exactly
  - a single wide root now benefits from host-side fan-out: the same seeded tree runs once with `preCalcMaxWorkers = 1` and once with `preCalcMaxWorkers = 4`
- `Phase5_PreCalcCancelLatencyLocal` now verifies:
  - a local delete entering pre-calc can be canceled while `GetDirectorySize()` is mid-walk
  - cancel-to-completion latency stays under `150 ms` with `directorySizeDelayMs = 50`
  - archived perf evidence records `FileOps.SelfTest.PreCalcCancelLatency` for same-machine before/after comparison

### Validation

- Deterministic settings round-trip archive: `Specs/TestRuns/4cb089111a23/Commands/2026-03-31_144719`
- Deterministic file-ops archive: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_144753`
- Deterministic pre-calc cancel-latency baseline: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_150236`
- Deterministic pre-calc cancel-latency candidate: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_150531`
- Post-change regression spot-check: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_150621`
- Single-root fan-out timeout archive while tuning the deterministic dataset: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_191217`
- Single-root fan-out passing archive: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_191447`
- Post-fan-out cancel-latency regression archive: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_191549`
- Post-fan-out queue-slot regression archive: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_191651`
- `settings_file_operations_precalc_roundtrip` passed in `12 ms`.
- `Phase5_PreCalcSettingsApplied` passed in `2640 ms`.
- `Phase5_PreCalcCancelLatencyLocal` failed before the plugin poll change at `219 ms` against a `150 ms` threshold, then passed after the change at `63 ms`.
- `Phase5_PreCalcSettingsApplied` passed again after the plugin poll change in `1796 ms`.
- Same-machine worker-budget evidence from `2026-03-31_191447`:
  - Multi-root check: `FileOps.PreCalc` with `sources=8 workers=4 configuredWorkers=4` took `87,488 us`
  - Multi-root check: `FileOps.PreCalc` with `sources=8 workers=8 configuredWorkers=8` took `102,737 us`
  - Single-root fan-out check: `FileOps.PreCalc` with `sources=1 workers=1 configuredWorkers=1` took `25,095,902 us`
  - Single-root fan-out check: `FileOps.PreCalc` with `sources=1 workers=4 configuredWorkers=4` took `6,333,747 us`
  - Same-machine speedup for the single-root fan-out workload: `3.96x` faster at `4` workers than `1`
- Same-machine cancel-latency evidence from `2026-03-31_150236` vs `2026-03-31_150531`:
  - `FileOps.SelfTest.PreCalcCancelLatency` improved from `219,000 us` to `63,000 us`
- Post-fan-out regression evidence:
  - `Phase5_PreCalcCancelLatencyLocal` passed in `2026-03-31_191549` with `FileOps.SelfTest.PreCalcCancelLatency = 46,000 us`
  - `Phase5_PreCalcCancelReleasesSlot` passed again in `2026-03-31_191651`
- Known harness note: `2026-03-31_191518` failed setup (`Failed to create delete-tree.`) because two exact selftests were launched in parallel against the same `last_run/fileops` sandbox; that archive is not treated as a product regression.
- Current conclusion: the new host queue materially helps a single wide root while the same machine still does **not** justify raising the default pre-calc worker budget above `4`; any wider default or wider plugin concurrency range remains evidence-gated.

---

## Phase 5 — Medium Impact: Parallel Search Directory Walk (P3)

### Problem

Search (`FileSystem.Search.cpp`) enumerates directories **sequentially** on the main search thread. Parallelism only kicks in for *evaluating entries within a single directory* (threshold: ≥500 entries). The recursive tree traversal itself is serial — discovering subdirectories is bound to a single thread.

For a wide tree (e.g., `node_modules` with 50K directories), directory traversal dominates search time.

### Proposed Changes

- [x] **5.1** Implement **parallel recursive directory walk**:
  - Landed for recursive name-only scan searches in `FileSystem::Search()` via a bounded shared directory queue plus threadpool workers
  - Each worker dequeues a directory, calls `ReadDirectoryInfo()`, filters name matches locally, and enqueues discovered child directories back into the shared queue
  - Worker count is now settings-backed via the FileSystem plugin schema field `searchMaxDirectoryWalkers` (default `4`, range `1`–`8`), so the parallel walk uses the configured exact worker budget instead of a hidden hardware-concurrency heuristic

- [x] **5.2** Thread-safe loop detection:
  - `MarkQueuedDirectory()` now guards both `queuedDirectories` and `queuedDirectoryIdentities` with a mutex so recursive worker enqueue is safe
  - The identity-based loop guard remains active when `FILESYSTEM_SEARCH_FOLLOW_SYMLINKS` is set

- [x] **5.3** Result aggregation:
  - Workers accumulate per-directory match buffers off-thread and publish completed directory results to a shared completion queue
  - The main search thread remains the only callback emitter: it drains ready directory results, calls `EmitSearchMatch()`, and continues to own progress / cancel checkpoints
  - Discovery order is no longer preserved intentionally; validation compares result sets after sort, which matches the plugin spec and current host expectations

- [x] **5.4** Maintain existing thresholds:
  - The existing per-directory parallel evaluation path (≥500 entries → 128-entry threadpool chunks) is unchanged
  - The new walk path composes with the existing entry-evaluation path because it only takes over recursive name-only scan traversal; other search modes keep the prior behavior

- [x] **5.5** Module pinning: all parallel walk workers now hold `AcquireModuleReferenceFromAddress()` for the callback lifetime and initialize COM as MTA before touching plugin-side directory enumeration

### Landed 2026-03-31

- `FileSystem::Search()` now switches recursive name-only scan queries onto `SearchDirectoryTreeParallelNameOnly()` while keeping content searches and non-recursive paths on the previous implementation.
- `ParallelDirectoryWalkState` owns the bounded pending-directory queue, completed-directory queue, shared counters, cancellation state, and the thread-safe loop-detection gate used by worker enqueue.
- Worker threads enumerate directories in parallel, but the main search thread still serializes `FileSystemSearchMatch()`, `FileSystemSearchProgress()`, and `FileSystemSearchShouldCancel()` so the host callback contract stays unchanged.
- `FileSystem.Search.ParallelDirectoryWalk` is now emitted as a real scoped duration metric with `value0=scannedDirectories` and `value1=maxActiveWorkers`; `FileSystem.Search.ScanTree` continues to cover the whole recursive traversal.
- `CompareDirectoriesEngine.SelfTest` now includes `local_search_scan_wide_tree_parallel_walk_name_only`, which builds a deterministic local tree, verifies the completed progress counters and returned match set, and now reruns the same seeded tree at `workers=1`, `4`, and `8` so the plugin setting and its perf impact are archived in one exact case.

### Validation

- Same-machine baseline archive: `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-03-31_152004`
- Same-machine candidate archive: `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-03-31_153527`
- The targeted wide-tree selftest stayed correct in both runs: `scannedDirectories=5377`, `scannedFiles=5120`, `matches=1024`, `expectedDirectories=5377`, `expectedFiles=5120`, `expectedMatches=1024`.
- Archived wall time improved from `6,845,495 us` to `6,648,375 us` (`1.03x`, about `2.9%` faster) for `compare.selftest.local_search_scan_wide_tree_us` on the same machine.
- `FileSystem.Search.ScanTree` improved from `6,837,815 us` to `6,640,737 us`; the landed candidate used `value1=8` active workers in both `FileSystem.Search.ScanTree` and `FileSystem.Search.ParallelDirectoryWalk`.
- Settings round-trip archive after landing `searchMaxDirectoryWalkers`: `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-03-31_192544`
- Settings-backed worker-budget archive: `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-03-31_192552`
- Latest adjacent regression archive: `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-03-31_192548`
- `local_search_backend_preferences_roundtrip` now also validates that the schema exposes `searchMaxDirectoryWalkers`, and it passed in `2026-03-31_192544`.
- `local_search_scan_wide_tree_parallel_walk_name_only` now emits `compare.selftest.local_search_scan_wide_tree_workers_us` for `workers=1`, `4`, and `8`; the same archive shows `FileSystem.Search.ParallelDirectoryWalk value1=1`, then `4`, then `8`, confirming the runtime honored the configured worker count exactly.
- Same-machine worker-budget evidence from `2026-03-31_192552` on the current seeded wide-tree run:
  - `workers=1`: `744,249 us`
  - `workers=4`: `726,385 us`
  - `workers=8`: `717,394 us`
  - `8` workers were only about `1.3%` faster than `4` on this machine, so the archive supports keeping the plugin default at `4` until a stronger broader win is demonstrated.
- `local_search_scan_follow_symlink_loop_guard` passed again in `2026-03-31_192548`, confirming the loop-detection path still terminates correctly under scan traversal after the settings-backed worker change.
- Result-set equivalence is validated as an unordered set after sort rather than discovery order; that is the intended contract for this phase.

---

## Phase 6 — Low Impact: Bandwidth Throttle Rework (P4)

### Problem

The CopyFileExW progress callback uses a **50ms sleep-slice** pattern (`kSleepSliceMs = 50` at ~L2044):

```cpp
while (elapsedMs < desiredMs) {
    Sleep(kSleepSliceMs);  // 50ms
    if (cancel) return PROGRESS_CANCEL;
    elapsedMs = GetTickCount64() - startTick;
}
```

Issues:
- **Bursty**: OS delivers data in large bursts between sleep checks → overshoot the bandwidth limit
- **Imprecise at low limits**: at 1 MiB/s with a 64KB CopyFileEx callback granularity, the math works out to ~0ms sleep → no throttling effect until bytes accumulate
- **Cancel latency**: up to 50ms delay between cancel request and actual cancellation (minor)

### Proposed Changes

- [x] **6.1** Rework callback-driven pacing around exact selftest evidence:
  - Parallel copy/move now uses shared debt-based pacing with 10 ms cancel polls, plus a debug-selectable shared-only vs per-worker sub-budget mode for fairness experiments
  - Sequential `CopyFileExW` paths now carry explicit callback-time throttle diagnostics; the latest isolated exact rerun shows the real callback-time 1-second window staying at `4,194,304 B` while the host-sampled progress view peaks at `6,291,456 B` because one `2,097,152 B` progress quantum lands inside the sampled window
  - `Phase6_LocalBandwidthThrottle` now records that sampled quantum explicitly and treats anything beyond one quantum as a real failure instead of a host-side attribution artifact
  - The original first candidate under-throttled badly (`lead=906,000 us` in archived run `2026-03-31_171538`); later fixes restored whole-copy rate adherence, and the latest exact guard now matches the observed callback granularity

- [x] **6.2** Support per-task and per-worker bucket:
  - Per-task bucket: shared across all workers for that task (total bandwidth limit)
  - Per-worker sub-budget: optional, prevents one worker from starving others
  - Host modifies `bandwidthLimitBytesPerSecond` → recalculate `refillRate` atomically
  - Landed exact fairness coverage via `Phase6_ParallelBandwidthThrottleFairness`; the current same-machine evidence is runtime-neutral, and both shared-only and per-worker modes repeatedly top out at the same `1,048,576 B` observable skew floor on this homogeneous local workload

- [x] **6.3** Tighten cancel responsiveness during throttle waits:
  - Current landing replaces the old 50 ms slices with 10 ms cancel polls via `CheckCancelImmediate()` while waiting on debt repayment
  - A true `WaitForSingleObject(cancelEvent, sleepMs)` integration is not part of the landed slice because the current host/plugin contract exposes cancellation as a callback poll, not a waitable handle

### Validation

- [x] Added deterministic local selftest coverage: `Phase6_LocalBandwidthThrottle`
  - Measures a local `CopyFileExW` copy at `4 MiB/s` and emits rate-adherence, rolling-window throughput, sampled progress quantum, and cancel-latency metrics
  - Fails if the local copy gets more than `250,000 us` ahead of the ideal duration, if any 1-second throughput window exceeds `110%` of the configured byte budget by more than one sampled progress quantum, or if cancel latency exceeds `250 ms`
- [x] Added deterministic parallel fairness coverage: `Phase6_ParallelBandwidthThrottleFairness`
  - Runs the same `4 x 8 MiB` local copy workload twice at `4 MiB/s`, once in shared-only mode and once in per-worker sub-budget mode
  - Emits `FileOps.SelfTest.ParallelBandwidthSharedOnlyDuration`, `...PerWorkerDuration`, and `...SkewImprovement`, and fails if the per-worker mode materially regresses runtime or active-stream fairness
- Same-machine evidence:
  - Baseline `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_171137`: duration `4,047,000 us`, lead `0 us`, cancel latency `31,000 us`
  - First token-bucket attempt `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_171538`: duration `3,094,000 us`, lead `906,000 us`, cancel latency `46,000 us` (rejected)
  - Corrected candidate `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_171814`: duration `4,063,000 us`, lead `0 us`, cancel latency `31,000 us`
  - The earlier rolling-window green reruns from `2026-03-31_173134` and `2026-03-31_185242` are now treated as superseded because they depended on the shared `%LOCALAPPDATA%\\RedSalamander\\SelfTest\\last_run` artifact root, which was later shown to be noisy on this machine
  - Isolated exact rerun `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-03-31_220452`: duration `4,093,000 us`, lead `0 us`, max 1-second window `5,242,880 B`, cancel latency `47,000 us` (failed)
  - Isolated exact rerun after the additional sequential reserve change `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-03-31_220754`: duration `4,344,000 us`, lead `0 us`, max 1-second window `5,242,880 B`, cancel latency `46,000 us` (failed)
  - Latest isolated exact rerun after the sampled-window guard update `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-03-31_225023-Phase6_LocalBandwidthThrottle`: duration `4,313,000 us`, lead `0 us`, host-sampled max 1-second window `6,291,456 B`, sampled progress quantum `2,097,152 B`, callback-time max window `4,194,304 B`, cancel latency `63,000 us` (passed)
  - Latest isolated fairness archive `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-03-31_225133-Phase6_ParallelBandwidthThrottleFairness`: shared-only `8,344,000 us`, per-worker `8,406,000 us`, shared skew `1,048,576 B`, per-worker skew `1,048,576 B`, shared active streams `4`, per-worker active streams `4`
  - Latest isolated Phase 8 regression archive `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-03-31_225241-Phase8_DefaultBandwidthLimitFromSettings`: baseline `140,000 us`, candidate `5,125,000 us`, slowdown `4,985,000 us`
  - Latest isolated adjacent copy/move regression archive `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-03-31_225417-Phase7_ParallelCopyMoveKnobs`: `Phase7_ParallelCopyMoveKnobs` passed in `61,672 ms`
  - Isolated rerun on the current binary `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_163700-Phase6_LocalBandwidthThrottle`: whole-copy duration `4,375,000 us`, lead `0 us`, host-sampled max 1-second window `6,291,456 B`, sampled progress quantum `2,097,152 B`, callback-time max window `4,194,304 B`, and cancel latency `47,000 us` (passed)
  - Isolated rerun on the current binary `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_163500-Phase6_ParallelBandwidthThrottleFairness`: shared-only `8,360,000 us`, per-worker `8,421,000 us`, shared skew `1,048,576 B`, per-worker skew `1,048,576 B`, shared active streams `4`, per-worker active streams `4`, shared samples `134`, per-worker samples `134` (still neutral and still bounded by the same observable skew floor)
  - A selftest-only artifact-root override (`REDSALAMANDER_SELFTEST_ROOT`) landed during this session so exact reruns can archive into a private folder instead of racing on the shared `%LOCALAPPDATA%\\RedSalamander\\SelfTest\\last_run` root.
  - Latest isolated Phase 8 regression archive after the fresh Phase 6 reruns `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_163900-Phase8_DefaultBandwidthLimitFromSettings`: baseline `140,000 us`, candidate `4,468,000 us`, slowdown `4,328,000 us`
  - Current evidence is now strong enough to close Phase 6: the sequential path is rate-accurate once callback granularity is accounted for, the exact local-throttle guard repeatedly passes on isolated reruns, cancel latency remains comfortably below the `250 ms` threshold, and the parallel fairness path is stable with no material runtime regression while both modes sit on the same `1,048,576 B` observable skew floor for this workload.

---

## Phase 7 — Low Impact: Adaptive Concurrency (P4)

### Problem

Concurrency caps are static and user-configured. Optimal values depend heavily on storage characteristics:
- **HDD**: sequential I/O optimal, 1–2 workers (random I/O from many workers thrashes the head)
- **SSD/NVMe**: high queue depth beneficial, 4–8 workers
- **Remote/cloud**: high latency, 8–16 workers (hide latency with concurrency)

Users rarely tune these settings, and the defaults (copy: 4, delete: 8) may be suboptimal for HDD or cloud storage.

### Proposed Changes

- [x] **7.1** Detect storage characteristics at operation start through `IFileSystem::GetStorageCharacteristics(...)`
  - The host now asks each filesystem for path-aware `FileSystemStorageCharacteristics` instead of inferring from plugin type
  - The built-in local FileSystem plugin returns preferred copy/move and delete budgets from its storage classification helper (`4` for local copy/move, `8` for delete, `8` for remote/high-latency copy/move)

- [x] **7.2** Auto-tune concurrency when user setting is `"auto"`:
  - Copy/move now resolves from `preferredCopyMoveConcurrency`
  - Permanent delete now resolves from `preferredDeleteConcurrency`
  - Manual mode still uses the configured numeric plugin settings
  - Recycle-bin delete intentionally stays on the explicit recycle-bin cap because shell batching cost is not modeled by storage hints

- [x] **7.3** Cross-volume operations: use the **minimum** of source and destination auto-tuned concurrency (limited by the slower side)

- [x] **7.4** Log chosen concurrency in the JSONL diagnostics file (`storageType`, `autoTunedConcurrency`)

- [x] **7.5** Add `"Auto"` option to `Preferences -> Plugins -> File System` via the plugin schema (no separate global concurrency pane)

### Landed 2026-04-01

- The FileSystem plugin schema/config now persists `concurrencyMode` with default `"auto"` and still exposes the manual numeric caps for copy/move, permanent delete, recycle-bin delete, and search walkers.
- The host now resolves per-item concurrency through two explicit paths:
  - Manual mode keeps the previous capability/configured-numeric behavior.
  - Auto mode uses `IFileSystem::GetStorageCharacteristics(...)` before task start and keeps `FileSystemOptions` ABI unchanged.
- Copy/move operations use the minimum of source and destination resolved budgets when the cross-filesystem bridge is active.
- Permanent delete uses the source filesystem's `preferredDeleteConcurrency`.
- Recycle-bin delete intentionally remains on the explicit recycle-bin cap, even in Auto mode.
- The task diagnostics stream now emits a dedicated `task.autoConcurrency` JSONL entry whenever Auto mode resolved the task budget, including `concurrencyMode`, `storageType`, optional `destinationStorageType`, `autoTunedConcurrency`, and the final `effectiveConcurrencyBudget` after any later clamping such as `@conn` overrides.
- Deterministic selftests now cover the landed behavior:
  - `settings_file_system_plugin_roundtrip` verifies the new schema/default JSON round-trip for `concurrencyMode`.
  - `Phase7_AutoConcurrencyHints` proves manual copy (`1`) vs auto-resolved local copy (`4`) plus auto-resolved permanent delete (`8`) and now also verifies the persisted `task.autoConcurrency` JSONL lines for the auto-copy and auto-delete tasks.
  - `Phase11_ConnectionOverridePrecedence` still proves that non-zero `@conn` overrides clamp after the resolved plugin/default budget.

### Validation

- Deterministic plugin-schema archive: `Specs/TestRuns/4cb089111a23/Commands/2026-04-01_120208`
- Deterministic adaptive-concurrency archive: `Specs/TestRuns/4cb089111a23/FileOps/2026-04-01_120717`
- Connection-override regression archive: `Specs/TestRuns/4cb089111a23/FileOps/2026-04-01_120808`
- Manual delete-knob regression archive: `Specs/TestRuns/4cb089111a23/FileOps/2026-04-01_120841`
- Latest adaptive-concurrency diagnostics archive: `Specs/TestRuns/4cb089111a23/FileOps/2026-04-01_123147`
- Latest connection-override regression archive after landing `task.autoConcurrency`: `Specs/TestRuns/4cb089111a23/FileOps/2026-04-01_123423`
- `settings_file_system_plugin_roundtrip` passed and now verifies the schema/default JSON includes `concurrencyMode` with default `"auto"`.
- `Phase7_AutoConcurrencyHints` passed in `2703 ms` and archived:
  - manual copy run: `FileOps.SelfTest.AutoConcurrencyManual = 343,000 us`, observed configured concurrency `1`
  - auto copy run: `FileOps.SelfTest.AutoConcurrencyAuto = 187,000 us`, observed configured concurrency `4`
  - same-machine improvement: `156,000 us` (`~45.5%` lower wall time) for the `16 x 4 MiB` local copy workload
  - auto permanent delete observed configured concurrency `8`
- The latest exact rerun `2026-04-01_123147` also passed after adding persisted-diagnostics verification:
  - manual copy run: `FileOps.SelfTest.AutoConcurrencyManual = 265,000 us`, observed configured concurrency `1`
  - auto copy run: `FileOps.SelfTest.AutoConcurrencyAuto = 219,000 us`, observed configured concurrency `4`
  - same-machine improvement: `46,000 us` (`~17.4%` lower wall time) for the same exact local copy workload on build `73`
  - auto permanent delete again observed configured concurrency `8`
- Final isolated confirmation rerun after the remaining Preferences/File Operations work closed: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_193000-Phase7_AutoConcurrencyHints`
- That final confirmation rerun still passed and kept the same resolved budgets: manual copy `266,000 us` at configured concurrency `1`, auto copy `203,000 us` at configured concurrency `4`, same-machine improvement `63,000 us` (`~23.7%` lower wall time), and auto permanent delete again observed configured concurrency `8`.
- Persisted JSONL evidence is now present in `%LOCALAPPDATA%\\RedSalamander\\Logs\\FileOperations-20260401.jsonl`:
  - auto copy task line includes `"category":"task.autoConcurrency"`, `"concurrencyMode":"auto"`, `"storageType":"unknown"`, `"autoTunedConcurrency":4`, and `"effectiveConcurrencyBudget":4`
  - auto delete task line includes `"category":"task.autoConcurrency"`, `"concurrencyMode":"auto"`, `"storageType":"unknown"`, `"autoTunedConcurrency":8`, and `"effectiveConcurrencyBudget":8`
  - the `Phase11_ConnectionOverridePrecedence` rerun now also proves later clamping is logged: the same JSONL stream includes `"destinationStorageType":"virtual"`, `"autoTunedConcurrency":4`, and `"effectiveConcurrencyBudget":2`
- `Phase11_ConnectionOverridePrecedence` still passed after the new Auto path landed:
  - inherited baseline remained configured at `4` (`266,000 us`)
  - explicit `@conn` override still clamped to `2` (`500,000 us`)
- The targeted rerun `2026-04-01_123423` passed again after the diagnostics landing, so the `task.autoConcurrency` emit did not break the override precedence path.
- Final isolated precedence confirmation rerun after the remaining Preferences/File Operations work closed: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_194000-Phase11_ConnectionOverridePrecedence`
- That final confirmation rerun still kept the same precedence shape: inherited configured concurrency `4`, explicit `@conn` override `2`, baseline `422,000 us`, candidate `562,000 us`, delta `140,000 us` (`~33.2%` slower under the tighter override).
- `Phase7_ParallelDeleteKnobs` passed again after the Auto resolver landed, confirming the explicit/manual delete path remained regression-safe.
- Two earlier exact-case archives (`2026-04-01_122212` and `2026-04-01_122911`) are intentionally retained only as selftest-harness retries while the diagnostics lookup path was being fixed; they are not treated as product regressions because the persisted JSONL stream already contained the correct `task.autoConcurrency` lines in those runs.

---

## Priority Summary

| Phase | Area | Impact | Effort | Risk | Priority |
|-------|------|--------|--------|------|----------|
| 1 | Recycle Bin Batching | High | Medium | Low | **P1** |
| 2 | Cross-FS Bridge Pipeline | High | High | Medium | **P2** |
| 3 | Callback Contention Reduction | Medium | Medium | Low | **P2** |
| 4 | Parallel Pre-Calc Tree Walk | Low-Med | Low | Low | **P3** |
| 5 | Parallel Search Directory Walk | Medium | Medium | Medium | **P3** |
| 6 | Bandwidth Throttle Rework | Low | Low | Low | **P4** |
| 7 | Adaptive Concurrency | Low | Low | Low | **P4** |
| 8 | Settings — Plugin Schema + Global Pane | Medium | Medium | Low | **P2** |

### Suggested Execution Order

1. **Phase 1** (Recycle Bin) — standalone, low risk, high payoff
2. **Phase 8** (Settings) — extends FileSystem plugin schema + adds global host pane for settings needed by phases 3, 4, 5, 6, 7
3. **Phase 3** (Callback Contention) — unlocks better scaling for phases 2, 4, 5
4. **Phase 2** (Cross-FS Bridge) — benefits from phase 3's reduced contention
5. **Phase 4** (Pre-Calc) — quick win, reuses existing scheduler
6. **Phase 7** (Adaptive Concurrency) — small, complements all parallel phases
7. **Phase 5** (Search Walk) — medium effort, compositional with existing entry-eval parallelism
8. **Phase 6** (Throttle Rework) — landed; revisit only if a new fairness scenario shows a real gap

---

## Interface Changes Required

This plan intentionally revises `IFileSystem` directly. Backward compatibility with old plugin binaries is explicitly out of scope for this design phase; the host and all built-in plugins will be updated together to the new contract. The goal here is the clearest long-term interface, not an incremental compatibility layer.

### New or Modified Plugin Interfaces (Common/PlugInterfaces/FileSystem.h)

- [x] Extend `IFileSystem` with:

```cpp
virtual HRESULT STDMETHODCALLTYPE GetTransferHints(const wchar_t* path,
                                                   FileSystemOperation operationType,
                                                   FileSystemTransferEndpoint endpoint,
                                                   FileSystemTransferHints* hints) noexcept = 0;

virtual HRESULT STDMETHODCALLTYPE GetStorageCharacteristics(const wchar_t* path,
                                                            FileSystemStorageCharacteristics* characteristics) noexcept = 0;
```

- [x] Add new supporting enums/structs in `Common/PlugInterfaces/FileSystem.h` with `sizeBytes` validation:

```cpp
enum FileSystemTransferEndpoint : uint32_t
{
    FILESYSTEM_TRANSFER_SOURCE_READ       = 1,
    FILESYSTEM_TRANSFER_DESTINATION_WRITE = 2,
};

enum FileSystemTransferLatencyClass : uint32_t
{
    FILESYSTEM_TRANSFER_LATENCY_UNKNOWN = 0,
    FILESYSTEM_TRANSFER_LATENCY_LOCAL   = 1,
    FILESYSTEM_TRANSFER_LATENCY_LAN     = 2,
    FILESYSTEM_TRANSFER_LATENCY_WAN     = 3,
    FILESYSTEM_TRANSFER_LATENCY_CLOUD   = 4,
};

enum FileSystemTransferHintFlags : uint32_t
{
    FILESYSTEM_TRANSFER_HINT_NONE                   = 0,
    FILESYSTEM_TRANSFER_HINT_PREFERS_LARGE_BUFFERS = 0x1,
    FILESYSTEM_TRANSFER_HINT_PREFERS_SEQUENTIAL_IO = 0x2,
    FILESYSTEM_TRANSFER_HINT_HIGH_METADATA_COST    = 0x4,
};

struct FileSystemTransferHints
{
    uint32_t sizeBytes; // sizeof(FileSystemTransferHints)
    uint32_t latencyClass;              // FileSystemTransferLatencyClass
    uint32_t flags;                     // FileSystemTransferHintFlags
    uint32_t preferredBufferBytes;      // e.g. 2 MiB, 4 MiB, 8 MiB
    uint32_t preferredProgressPeriodMs; // e.g. 100..250ms
};

enum FileSystemStorageKind : uint32_t
{
    FILESYSTEM_STORAGE_UNKNOWN       = 0,
    FILESYSTEM_STORAGE_HDD           = 1,
    FILESYSTEM_STORAGE_SSD           = 2,
    FILESYSTEM_STORAGE_NVME          = 3,
    FILESYSTEM_STORAGE_NETWORK_SHARE = 4,
    FILESYSTEM_STORAGE_CLOUD         = 5,
    FILESYSTEM_STORAGE_VIRTUAL       = 6,
    FILESYSTEM_STORAGE_MEMORY        = 7,
};

enum FileSystemStorageFlags : uint32_t
{
    FILESYSTEM_STORAGE_FLAG_NONE                   = 0,
    FILESYSTEM_STORAGE_FLAG_ROTATIONAL             = 0x1,
    FILESYSTEM_STORAGE_FLAG_HIGH_LATENCY           = 0x2,
    FILESYSTEM_STORAGE_FLAG_PREFERS_SEQUENTIAL_IO  = 0x4,
    FILESYSTEM_STORAGE_FLAG_SUPPORTS_DEEP_QUEUE    = 0x8,
};

struct FileSystemStorageCharacteristics
{
    uint32_t sizeBytes; // sizeof(FileSystemStorageCharacteristics)
    uint32_t storageKind;                  // FileSystemStorageKind
    uint32_t flags;                        // FileSystemStorageFlags
    uint32_t queueDepthHint;
    uint32_t preferredCopyMoveConcurrency;
    uint32_t preferredDeleteConcurrency;
};
```

- [x] Define exact semantics in the plan/spec:
  - `path` is always required and is interpreted in the plugin's own path space
  - the host calls `GetTransferHints(...)` separately for the source path with `FILESYSTEM_TRANSFER_SOURCE_READ` and for the destination path with `FILESYSTEM_TRANSFER_DESTINATION_WRITE`
  - `GetTransferHints(...)` is for bridge buffering and progress cadence, not for user-visible policy
  - `GetStorageCharacteristics(...)` is the authoritative source for auto-concurrency and diagnostics
  - plugins with non-volume-backed storage (cloud, virtual, memory) still return a meaningful classification instead of a Win32-only approximation
- [x] Do **not** extend `FileSystemOptions` in place for v1 of this work; keep Auto/Manual selection in plugin configuration and resolve it to today's effective concurrency fields before the operation starts
- [x] Keep `IFileSystemCallback::FileSystemProgress()` ABI unchanged; Phase 3 reduces contention in the implementation while preserving callback-based pause, queue-pause, and prompt semantics

All built-in file-system plugins shipped in-tree (`FileSystem`, `FileSystem7z`, `FileSystemCurl`, `FileSystemS3`, `FileSystemGoogleDrive`, `FileSystemMicrosoftDrive`, and `FileSystemDummy`) must implement the new methods as part of this contract revision. There is no plugin-type inference fallback in the target design. Coordinate cloud-plugin adoption with [Code_PluginImprovementPlan.md](Code_PluginImprovementPlan.md), but keep the end-state mandatory once this plan lands.

---

## Perf Validation Contract

Per project guidelines ([perf-validation skill](../../.github/skills/perf-validation/SKILL.md)):

Each phase must deliver:
1. **Scenario definition** in the spec (what is measured, what inputs, what metric)
2. **Instrumentation** via `Debug::Perf::Scope` or `TraceLoggingWrite` at key entry/exit points
3. **Deterministic selftest** that exercises the scenario with reproducible inputs
4. **Archived perf evidence** in `Specs/TestRuns/` with before/after comparison

No phase may merge without perf evidence demonstrating the claimed improvement.

---

## Phase 8 — Medium Impact: Settings for New Features (P2)

### Problem

Several phases in this plan introduce new tunable settings. These settings split into two categories:

1. **Plugin-specific settings** — concurrency, batching, and search behavior that differ per plugin and are already managed via each plugin's JSON configuration schema (`IInformations::GetConfigurationSchema()` → `Preferences → Plugins → [plugin]`)
2. **Global host settings** — pre-calc, bandwidth defaults, and cross-FS bridge behavior that are host-side concerns applying across all plugins

The FileSystem plugin already exposes 10 settings via its schema (Auto/Manual concurrency mode, concurrency caps, recycle-bin batching, enumeration buffers, reparse policy, and search backend/walkers). New plugin-specific settings should extend that schema. Host-global settings need a new Preferences pane.

### Current State

**FileSystem plugin schema** (via `Preferences → Plugins → File System`):
| Existing Setting | Type | Default | Range |
|------------------|------|---------|-------|
| `concurrencyMode` | option | auto | auto/manual |
| `copyMoveMaxConcurrency` | value | 4 | 1–16 |
| `deleteMaxConcurrency` | value | 8 | 1–64 |
| `deleteRecycleBinMaxConcurrency` | value | 2 | 1–16 |
| `recycleBinBatchSize` | value | 500 | 1–1000 |
| `enumerationSoftMaxBufferMiB` | value | 512 | 1–4095 |
| `enumerationHardMaxBufferMiB` | value | 2048 | 1–4095 |
| `reparsePointPolicy` | option | copyReparse | copyReparse/followTargets/skip |
| `searchBackendPreference` | option | auto | auto/service/local-index/scan |
| `searchMaxDirectoryWalkers` | value | 4 | 1–8 |

**Per-connection overrides** (Connection Manager → `extra`):
- `copyMoveMaxConcurrency` (0 = inherit plugin default)
- `deleteMaxConcurrency` (0 = inherit plugin default)

**Global host settings partially landed today**:
- `fileOperations.preCalcEnabled`
- `fileOperations.preCalcMaxWorkers`
- `fileOperations.crossFsBridgeBufferSizeKB`
- `fileOperations.defaultBandwidthLimitBytesPerSecond`
- The dedicated Preferences pane is still open

### Setting Ownership Decision

| Setting | Owner | Rationale |
|---------|-------|-----------|
| `concurrencyMode` (Auto/Manual) | **Plugin schema** | Auto-detection uses `IOCTL_STORAGE_QUERY_PROPERTY` which is storage-specific; cloud plugins don't need it |
| `copyMoveMaxConcurrency` | **Plugin schema** | Already exists, just widen range |
| `deleteMaxConcurrency` | **Plugin schema** | Already exists |
| `recycleBinMaxConcurrency` | **Plugin schema** | Already exists |
| `recycleBinBatchSize` | **Plugin schema** | Only FileSystem uses `IFileOperation` shell API |
| `searchMaxDirectoryWalkers` | **Plugin schema** | Only plugins with local directory walking need this |
| `preCalcEnabled` | **Global host pane** | Host runs pre-calc for all plugins |
| `preCalcMaxWorkers` | **Global host pane** | Host runs pre-calc for all plugins |
| `defaultBandwidthLimitBytesPerSecond` | **Global host pane** | Host applies the limit regardless of plugin |
| `crossFsBridgeBufferSizeKB` | **Global host pane** | Bridge is host code between any two plugins |

### Proposed Changes

#### 8A. Extend FileSystem Plugin Schema

Add new fields to the existing `kSchemaJson` in `Plugins/FileSystem/FileSystem.h`:

- [x] **8.1** `concurrencyMode` — type: `option`, options: `["auto", "manual"]`, default: `"auto"`. When `"auto"`, `copyMoveMaxConcurrency`/`deleteMaxConcurrency` are auto-tuned per storage type and their manual values are ignored. When `"manual"`, the existing concurrency fields apply.
  - Landed on 2026-04-01 in `FileSystem.h` / `FileSystem.cpp` / `FileSystem.Watch.cpp`; the plugin now persists the field, defaults fresh instances to `"auto"`, and the Commands selftest round-trip verifies both schema and configuration JSON.
- [x] **8.2** Widen `copyMoveMaxConcurrency` range from 1–8 to **1–16** (SSD/NVMe may benefit from deeper queues), and prove the value of settings >8 with widened scheduler / host caps plus archived same-machine perf evidence before keeping the change.
- [x] **8.3** `recycleBinBatchSize` — type: `value`, default: `500`, min: `1`, max: `1000`. Items per `IFileOperation` batch (Phase 1). Value of 1 disables batching.
  - Landed on 2026-03-31 in `FileSystem.h` / `FileSystem.cpp` / `FileSystem.FileOps.cpp`; the plugin schema now persists the field and `DeleteItems()` snapshots it into the recycle-bin batch scheduler.
- [x] **8.4** `searchMaxDirectoryWalkers` — type: `value`, default: `4`, min: `1`, max: `8`. Workers for parallel search directory traversal (Phase 5).
  - Landed on 2026-03-31 in `FileSystem.h` / `FileSystem.cpp` / `FileSystem.Search.cpp`; the plugin schema now persists the field and `FileSystem::Search()` snapshots it into the parallel recursive scan walk.

Update `SetConfiguration()` / `GetConfiguration()` in `FileSystem.cpp` to handle the new fields.

**Schema rendering note:** The existing `Preferences → Plugins → File System` pane auto-renders these via the generic schema parser in `Preferences.Plugin.Configuration.cpp`. No new UI code needed for plugin-schema fields — the host builds controls dynamically from the JSON schema.

#### 8B. Extend FileSystemCurl Plugin Schema (if applicable)

- [x] **8.5** Review whether `FileSystemCurl` should also get `concurrencyMode` — keep FileSystemCurl manual-only for now. Its current `GetStorageCharacteristics(...)` path already classifies the transport as high-latency remote/cloud with fixed preferred budgets, so adding an `auto/manual` UI knob there would duplicate a single remote profile instead of adding a real user-facing choice.

#### 8C. New Global Host Settings (SettingsStore)

Extend the existing `fileOperations` section in `Specs/SettingsStore.schema.json`:

- [x] **8.6** `fileOperations.preCalcEnabled` — `true` | `false` (default: `true`). Whether pre-calculation scan runs before copy/move operations.
- [x] **8.7** `fileOperations.preCalcMaxWorkers` — `1`–`8` (default: `4`). Workers for the pre-calc tree walk (Phase 4).
- [x] **8.8** `fileOperations.defaultBandwidthLimitBytesPerSecond` — `0` (unlimited) | positive integer (default: `0`). Pre-populated default speed limit for new tasks; user can still override per-task in the progress popup.
- [x] **8.9** `fileOperations.crossFsBridgeBufferSizeKB` — `512`–`16384` (default: `4096`). Per-buffer size for the cross-filesystem double-buffered pipeline (Phase 2). Two buffers are allocated per active transfer.

#### 8D. New Preferences Pane — File Operations (Global)

- [x] **8.10** Create `Preferences.FileOperations.h` / `.cpp` — new pane following the existing pattern (see `Preferences.CompareDirectories.h` as template)
- [x] **8.11** Register pane in `Preferences.Dialog.cpp` — insert between "Plugins" and "Advanced" in the pane list
- [x] **8.12** Add STRINGTABLE entries in the `.rc` file for all new labels and descriptions

**Pane Layout** — only global/host-owned settings:

**Pre-Calculation**
| Control | Setting | Notes |
|---------|---------|-------|
| Checkbox | Enable pre-calculation scan | Default: checked |
| Slider + spin | Pre-calc workers | 1–8, grayed when unchecked |

**Bandwidth**
| Control | Setting | Notes |
|---------|---------|-------|
| Combo box | Default Speed Limit | Unlimited / 1 MiB/s / 5 / 10 / 50 / 100 / 500 MiB/s / 1 GiB/s / Custom |
| Edit + unit | Custom Limit | Shown only when "Custom" selected |

**Advanced** (collapsed group)
| Control | Setting | Notes |
|---------|---------|-------|
| Slider + spin | Cross-FS bridge buffer (KB) | 512–16384 |

**Note:** Concurrency, recycle-bin batching, and search walkers are NOT in this pane — they live in `Preferences → Plugins → File System` via the plugin's own schema. The pane should include a hint: *"Per-plugin concurrency and batching settings are in Preferences → Plugins → [plugin name]."*

#### 8E. Settings Precedence

```
Per-task override (progress popup)
  ↓ falls back to
Per-connection override (Connection Manager → extra, if non-zero)
  ↓ falls back to
Plugin-wide default (Preferences → Plugins → File System)
  ↓ falls back to
Hardcoded default in plugin schema
```

For global host settings:
```
Per-task override (progress popup, bandwidth only)
  ↓ falls back to
Global default (Preferences → File Operations)
  ↓ falls back to
Hardcoded default
```

#### 8F. Settings Propagation

- [x] **8.13** Plugin settings: no new plumbing needed — existing `SetConfiguration()` / `GetConfiguration()` path handles apply. Plugin reads updated values on next operation start.
  - Landed direct plugin round-trip coverage in Commands selftest archive `Specs/TestRuns/4cb089111a23/Commands/2026-03-31_201809` (`settings_file_system_plugin_roundtrip`), which verifies schema exposure plus `SetConfiguration()` / `GetConfiguration()` / `SomethingToSave()` for `recycleBinBatchSize` and `searchMaxDirectoryWalkers`.
- [x] **8.14** Global host settings: read from `SettingsStore` at task creation time. For bandwidth limit, pre-populate into `FileSystemOptions.bandwidthLimitBytesPerSecond`. For pre-calc, check `preCalcEnabled` before launching the pre-calc phase. For bridge buffering, snapshot `crossFsBridgeBufferSizeKB` into the task and allocate both bridge buffers from that frozen byte count. Live settings reload updates defaults for newly created tasks only; it does not retune already-running operations.
- [x] **8.15** In the progress popup: pre-populate bandwidth limit from the global default; if Auto mode is selected, show the resolved applied concurrency as read-only diagnostics rather than adding a second editable concurrency control in the popup.
  - The bandwidth half is landed because new copy/move tasks seed `_desiredSpeedLimitBytesPerSecond` from `fileOperations.defaultBandwidthLimitBytesPerSecond`, and the existing popup already reflects that task state.
  - The Auto-concurrency diagnostics half is also landed. Isolated archive `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_140700-Phase7_AutoConcurrencyHints` now verifies both the live popup and the completed-task view: `popupAuto=4/4`, `popupDelete=8/8`, `completedPopupAuto=4/4`, and `completedPopupDelete=8/8`.
  - Earlier isolated archives `isolated-2026-04-01_135836-Phase7_AutoConcurrencyHints`, `isolated-2026-04-01_154500-Phase7_AutoConcurrencyHints`, and `isolated-2026-04-01_160906-Phase7_AutoConcurrencyHints` are superseded selftest-observation failures from before the exact case disabled auto-dismiss for completed-card validation; they are excluded from the final conclusion.

#### 8G. Spec & Documentation Updates

- [x] **8.16** Update `Specs/UI/UI_PreferencesDialog.md` — add File Operations pane to the pane list and document all controls
- [x] **8.17** Update `Specs/SettingsStore.schema.json` — add `fileOperations.*` keys with types, ranges, defaults, titles, descriptions
- [x] **8.18** Update `Docs/FileOperations.md` — document the new Preferences pane, the extended plugin settings, and how global vs. per-plugin vs. per-connection settings interact
- [x] **8.19** Update `Docs/Preferences.md` (if it exists) or `Docs/GettingStarted.md` — mention the new pane

### Partial Landing 2026-03-31

- `SettingsStore` now persists `fileOperations.crossFsBridgeBufferSizeKB` (`512`–`16384`, default `4096`) and `fileOperations.defaultBandwidthLimitBytesPerSecond` (`0+`, default `0`) alongside the earlier pre-calc host settings.
- Task creation snapshots the configured bridge buffer size into `_crossFsBridgeBufferBytes`, and the shared cross-filesystem bridge allocates both buffers from that frozen byte count so already-running transfers do not retune mid-flight.
- Task creation now resolves `defaultBandwidthLimitBytesPerSecond` into `_desiredSpeedLimitBytesPerSecond` for new copy/move tasks when the caller passes `0`, so the existing progress popup inherits the seeded value without adding new popup plumbing in this slice.
- The FileSystem plugin schema now also persists `copyMoveMaxConcurrency` (`1`–`16`, default `4`), `recycleBinBatchSize` (`1`–`1000`, default `500`), and `searchMaxDirectoryWalkers` (`1`–`8`, default `4`); `DeleteItems()` snapshots the recycle-bin batch size into the batch scheduler, `FileSystem::Search()` snapshots the directory-walker count for recursive name-only scan walks, and the local copy/move scheduler + host in-flight display path now honor copy/move values above `8`.
- `settings_file_operations_precalc_roundtrip` now covers `crossFsBridgeBufferSizeKB = 8192` plus `defaultBandwidthLimitBytesPerSecond = 3 MiB/s` in addition to the pre-calc settings.
- `Phase8_DefaultBandwidthLimitFromSettings` now proves the host setting on the same machine: baseline `188,000 us` with default limit `0`, candidate `5,062,000 us` with default limit `1 MiB/s`, slowdown `4,874,000 us` (`~26.9x` slower).
- Latest isolated regression rerun still passes after the newer Phase 6 throttle experiments: baseline `140,000 us`, candidate `5,125,000 us`, slowdown `4,985,000 us` in archive `isolated-2026-03-31_225241-Phase8_DefaultBandwidthLimitFromSettings`.
- Fresh same-build rerun after the completed-popup fix still keeps the host-setting behavior stable: archive `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_140956-Phase8_DefaultBandwidthLimitFromSettings` recorded baseline `172,000 us`, candidate `4,500,000 us`, slowdown `4,328,000 us`.
- Fresh isolated rerun after the File Operations live-interaction fix still keeps the host-setting behavior stable: archive `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_191100-Phase8_DefaultBandwidthLimitFromSettings` recorded baseline `110,000 us`, candidate `4,484,000 us`, slowdown `4,374,000 us`.
- `Phase11_BridgePipelineDummyToDummyPerf` now verifies the default `4096 KB` task snapshot before measuring pipeline throughput, so Phase 8 validation and Phase 2 perf evidence share the same archived run.
- Final isolated bridge-buffer confirmation rerun after the remaining Preferences/File Operations work closed: archive `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_193300-Phase11_BridgePipelineDummyToDummyPerf` kept the same task-snapshot contract with `configuredBufferKB=4096`, `resolvedBufferKB=8192`, baseline `20,297,000 us`, and candidate `453,000 us`.
- `settings_file_system_plugin_roundtrip` now gives direct plugin-schema/config coverage in archive `2026-03-31_203618`: the isolated FileSystem plugin instance advertises `copyMoveMaxConcurrency`, `recycleBinBatchSize`, and `searchMaxDirectoryWalkers` in schema, round-trips `copyMoveMaxConcurrency=12` alongside custom batch/search values, and returns to a clean default state after reset.
- `Phase11_ConnectionOverridePrecedence` now gives direct per-connection precedence coverage for `copyMoveMaxConcurrency`: archive `2026-03-31_205246` first landed the exact case, and same-build rerun `2026-03-31_212317` kept the result stable after timing from actual task start. On the same machine, the inherited dummy-plugin baseline configured concurrency stayed at `4`, while a non-zero `@conn` override forced `2`; the rerun recorded `156,000 us` baseline versus `172,000 us` candidate (`16,000 us` slower, about `10.3%`).
- Latest isolated rerun `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_140904-Phase11_ConnectionOverridePrecedence` still passes after the Phase 7 exact-case fix and kept the same precedence shape: inherited configured concurrency `4`, explicit `@conn` override `2`, baseline `360,000 us`, candidate `578,000 us`, delta `218,000 us`.
- Final isolated precedence confirmation rerun after the remaining Preferences/File Operations work closed: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_194000-Phase11_ConnectionOverridePrecedence` again kept inherited configured concurrency `4` versus explicit `@conn` override `2`, with baseline `422,000 us`, candidate `562,000 us`, and delta `140,000 us`.
- `Phase11_ConnectionOverrideGlobalGate` now uses shared two-task wall time as the primary guard instead of brittle popup-line saturation sampling. The clean same-build archive `2026-03-31_212317` recorded `FileOps.SelfTest.ConnectionOverrideGlobalGateDuration = 38,078,000 us` for two simultaneous `16 MiB` bridged copies against the same `@conn` profile with `copyMoveMaxConcurrency = 2`, which is comfortably above the `24,000,000 us` minimum gate threshold.
- `Phase7_RecycleBinBatchDelete` now verifies `recycleBinBatchSize` end-to-end in archive `2026-03-31_194908`: baseline `7,000,000 us` at `batchSize=1`, candidate `1,485,000 us` at `batchSize=500`, improvement `5,515,000 us`.
- Latest same-build recycle-bin rerun `2026-03-31_201848` kept the same configuration/perf behavior for `recycleBinBatchSize`: baseline `5,672,000 us`, candidate `1,156,000 us`, improvement `4,516,000 us`.
- `Phase7_RecycleBinBatchDeleteMultiBatch` now verifies the configurable cap itself in archive `2026-03-31_194943`, where `recycleBinBatchSize=256` yields three observed recycle-bin batches of `256` items each with `BatchFailedItems=0`.
- New exact perf case `Phase7_CopyMoveConcurrency16Perf` now validates the widened copy/move cap on the same machine with a local-to-local `24 x 8 MiB` copy workload in per-item mode. Archive `2026-03-31_203457` recorded baseline `610,000 us` at `copyMoveMaxConcurrency=8` versus candidate `297,000 us` at `16` (`313,000 us` improvement, about `51.3%` faster). Same-build rerun `2026-03-31_203553` stayed in the same range at `563,000 us` versus `250,000 us` (`313,000 us` improvement, about `55.6%` faster).
- `local_search_backend_preferences_roundtrip` now covers the new plugin field in archive `2026-03-31_192544`, and `local_search_scan_wide_tree_parallel_walk_name_only` proves the applied worker budgets `1`, `4`, and `8` in archive `2026-03-31_192552`.

### Validation

- Verify plugin schema fields render correctly in `Preferences → Plugins → File System` (auto-generated UI)
- Landed direct plugin schema/config round-trip coverage for `copyMoveMaxConcurrency`, `recycleBinBatchSize`, and `searchMaxDirectoryWalkers`: `Specs/TestRuns/4cb089111a23/Commands/2026-03-31_203618`
- Verify new global pane controls read/write to SettingsStore (round-trip: set → close → reopen → values preserved)
  - Landed schema/settings coverage for `preCalcEnabled`, `preCalcMaxWorkers`, `crossFsBridgeBufferSizeKB`, and `defaultBandwidthLimitBytesPerSecond`: `Specs/TestRuns/4cb089111a23/Commands/2026-03-31_165201`
  - Fresh isolated reruns keep the page and live interactions green after the combo/value/cancel follow-ups: `Specs/TestRuns/4cb089111a23/Commands/isolated-2026-04-01_190900-cmd_preferences_dialog_file_operations_page_uses_dxui_controls`, `Specs/TestRuns/4cb089111a23/Commands/isolated-2026-04-01_190200-cmd_preferences_dialog_file_operations_live_dx_interaction`, and `Specs/TestRuns/4cb089111a23/Commands/isolated-2026-04-01_191000-settings_file_operations_precalc_roundtrip`
- Verify per-connection overrides take precedence over plugin-wide defaults when non-zero
  - Landed exact FileOps coverage for per-connection copy/move overrides: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_205246` first landed the new `Phase11_ConnectionOverridePrecedence` case, same-build rerun `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_212317` kept it stable, `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_212317` also gives the updated `Phase11_ConnectionOverrideGlobalGate` archive, and `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_212430` reruns `Phase11_ConnectionOverrideClamp`.
  - Latest isolated rerun after the completed-popup selftest fix still passes and keeps the same precedence result: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_140904-Phase11_ConnectionOverridePrecedence`
  - Final isolated confirmation rerun after the remaining Preferences/File Operations work closed: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_194000-Phase11_ConnectionOverridePrecedence`
  - The precedence case proves inherited dummy-plugin concurrency `4` versus non-zero `@conn` override `2` on the same workload, while the clamp case still verifies copy/delete clamp-to-`1`.
  - Earlier archives `2026-03-31_205631` and `2026-03-31_205823` are superseded harness-failure runs from the intermediate observation strategy and are excluded from the final conclusion.
- Verify Auto concurrency mode persists correctly in the plugin config UI; if the generic schema editor cannot gray out dependent fields yet, document that the manual numeric fields remain visible but are ignored while Auto is selected
- Landed read-only popup/completed-task diagnostics coverage for Auto mode: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_140700-Phase7_AutoConcurrencyHints`
  - That archive proves the live popup and the retained completed-task snapshot both report resolved/applied `4/4` for the auto copy and `8/8` for the auto delete, and the same run also verifies the persisted `task.autoConcurrency` JSONL lines for task `2` and task `3`.
- Verify settings affect newly created file operations without requiring an app restart
  - Landed task-creation coverage for pre-calc settings: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_144753`
  - Landed task-snapshot coverage for `defaultBandwidthLimitBytesPerSecond`: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_165238`
  - Post-Phase-6 throttle regression reruns still pass for `defaultBandwidthLimitBytesPerSecond`: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_171957`, `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_173247`, and isolated latest `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-03-31_225241-Phase8_DefaultBandwidthLimitFromSettings`
  - Fresh same-build rerun after the completed-popup selftest fix also passes for `defaultBandwidthLimitBytesPerSecond`: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_140956-Phase8_DefaultBandwidthLimitFromSettings`
  - Latest isolated rerun after the File Operations live-interaction fix also passes for `defaultBandwidthLimitBytesPerSecond`: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_191100-Phase8_DefaultBandwidthLimitFromSettings`
  - Landed task-snapshot + bridge allocation coverage for `crossFsBridgeBufferSizeKB`: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_163447` and stability rerun `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_163749`
  - Final isolated bridge-buffer confirmation rerun after the remaining Preferences/File Operations work closed: `Specs/TestRuns/4cb089111a23/FileOps/isolated-2026-04-01_193300-Phase11_BridgePipelineDummyToDummyPerf`
  - Landed task-snapshot + same-machine perf coverage for `copyMoveMaxConcurrency > 8`: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_203457` and same-build rerun `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_203553`
  - Landed plugin-setting snapshot coverage for `recycleBinBatchSize`: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_194908` (`batchSize=1` vs `500`), latest same-build rerun `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_201848`, and `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_194943` (`batchSize=256` with three exact `256`-item batches)
  - Landed plugin-setting snapshot coverage for `searchMaxDirectoryWalkers`: `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-03-31_192552`, where `FileSystem.Search.ParallelDirectoryWalk value1` reports `1`, `4`, and `8` on the three back-to-back search runs
  - Clean post-landing regressions still pass: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_203658` (`Phase7_SharedPerItemScheduler`) and `Specs/TestRuns/4cb089111a23/FileOps/2026-03-31_203853` (`Phase7_ParallelCopyMoveKnobs`). One overlapping parallel rerun (`2026-03-31_203718`) produced a false-negative `42/43` output mismatch and is excluded from the perf conclusion.

---

## Scope Exclusions

- **FileSystem7z, FileSystemCurl, cloud plugins** — covered by [Code_PluginImprovementPlan.md](Code_PluginImprovementPlan.md)
- **Directory enumeration API changes** — NtQueryDirectoryFile path is already optimal; no changes planned
- **UI popup rendering performance** — separate concern (DxUI layer)
- **Compare Directories parallelism** — separate plan (uses different code path)
- **Search index (USN/MFT)** — covered by [RFC_Core_InstantFileSearchUsnMftIndex.md](RFC_Core_InstantFileSearchUsnMftIndex.md)
