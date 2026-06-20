# File Operations Engine (Developer)

The File Operations Engine runs every long-running Copy/Move/Delete off the UI thread, gates them through a Wait/Parallel queue, dispatches per-item work on a shared scheduler, arbitrates conflicts inline, and moves bytes between unrelated providers through a host-driven cross-filesystem bridge. This page documents the internals; for the user-facing behavior see [../FileOperations.md](../FileOperations.md), for the higher-level subsystem overview see [../DeveloperGuide.md](../DeveloperGuide.md), and for the normative contract see [Specs/FileSystem/FileSystem_FileOperations.md](../../Specs/FileSystem/FileSystem_FileOperations.md).

## Source layout

The engine is split across one header and several translation units so a single .cpp does not grow unbounded. All symbols live under `FolderWindow::FileOperationState` and its nested `Task`.

| File | Role |
|------|------|
| `RedSalamander/FolderWindow.FileOperationsInternal.h` | Private header: `FileOperationState`, nested `Task` (implements `IFileSystemCallback` + `IFileSystemDirectorySizeCallback`), `ConflictArbiter`, `ConflictBucket`/`ConflictAction`, `PerItemCallbackCookie`, `PerfStats`, `CompletedTaskSummary`, `TaskDiagnosticEntry` |
| `RedSalamander/FolderWindow.FileOperations.State.Runtime.Part.cpp` | `StartOperation`, `ApplyQueueMode`, `ShouldQueueNewTask`, `HasActiveOperations`, `CancelAll`, `NotifyQueueChanged`, `Shutdown` |
| `RedSalamander/FolderWindow.FileOperations.State.Queue.Part.cpp` | `EnterOperation`, `LeaveOperation`, `PostCompleted`, `RemoveFromQueue`, `UpdateQueuePausedTasks`, `FindTask`, `RemoveTask`, plus `ENABLE_TESTS` hooks |
| `RedSalamander/FolderWindow.FileOperations.State.cpp` | `Task::ThreadMain`, `ExecuteOperation`, `RunPreCalculation`, the `IFileSystemCallback` methods, the conflict arbiter helpers, `PerItemTaskScheduler`, and the `CrossFileSystemBridge` |
| `RedSalamander/FolderWindow.FileOperations.State.Diagnostics.Part.cpp` | `RecordCompletedTask`, diagnostics log + Issues capture |
| `RedSalamander/FolderWindow.FileOperations.Popup.cpp` / `.IssuesPane.cpp` | Direct2D/DirectWrite UI; reads published snapshots, posts user actions back via `SubmitConflictDecision` etc. |

## Per-task jthread model

`StartOperation` is the single entry point. It validates preconditions (provider capabilities through `CanSameFileSystemOperation`, same-folder / overlap rejection, destructive-delete confirmation), snapshots host-owned settings into the new `Task` (`_enablePreCalc`, `_preCalcMaxWorkers`, `_crossFsBridgeBufferBytes`, the seeded `_desiredSpeedLimitBytesPerSecond`), pushes the `unique_ptr<Task>` into `_tasks` under `_mutex`, ensures the popup is visible, and then spawns exactly one worker:

```cpp
rawTask->_thread = std::jthread([rawTask](std::stop_token stopToken) noexcept { rawTask->ThreadMain(stopToken); });
```

There is one `std::jthread` per task (`Task::_thread`). The `std::stop_token` is the cooperative cancellation channel: `ThreadMain` installs a `std::stop_callback` that, on stop request, wakes `_pauseCv`, the `_conflictArbiter.cv`, and the state's `_queueCv` so a task blocked anywhere unblocks promptly. Cancellation itself is requested via `Task::RequestCancel` (sets `_cancelled`); both `_cancelled` and the stop token are polled at every checkpoint.

`ThreadMain` is the control-flow spine:

1. `SetWaitingInQueue(_waitForOthers)` so the UI shows "Waiting" immediately for queued tasks.
2. `_state->EnterOperation(*this, stopToken)` — the start gate (below). If it returns `false` the task was cancelled while queued; it stores `ERROR_CANCELLED` and calls `PostCompleted`.
3. On success it latches `_enteredOperationTick` / `_enteredOperation` (used by queue-pause ordering).
4. Runs pre-calc (concurrent or serial, see "5F early admission").
5. `ExecuteOperation()` — the actual transfer/delete.
6. Joins the side pre-calc thread (if any), records perf (`FileOps.Operation`, `FileOps.PreCalc`, `FileOps.Bridge.*`), logs the result diagnostic, and falls through to teardown. `PostCompleted` posts `WndMsg::kFileOperationCompleted` to the FolderWindow.

On `Shutdown` the state cancels all live tasks (`RequestCancel`) and lets the `_tasks` vector destruct; each `jthread` joins automatically.

## EnterOperation: the start-gating queue

`EnterOperation` (in `.Queue.Part.cpp`) implements Wait vs Parallel start gating against `FileOperationState`-owned state: `_queueMutex`, `_queueCv`, a FIFO `std::deque<uint64_t> _queue` of waiting task IDs, and an `unsigned long _activeOperations` count.

- If the task's `_waitForOthers` is false (Parallel), it simply `++_activeOperations` and returns `true` — no blocking.
- If `_waitForOthers` is true (Wait), it pushes its task ID onto `_queue`, calls `NotifyQueueChanged`, and waits on `_queueCv` until one of:
  - the stop token fires or `_cancelled` is set (returns `false` after `RemoveFromQueue`), or
  - `_waitForOthers` flipped to false mid-wait (mode switched to Parallel) — it removes itself from the queue and proceeds, or
  - `_activeOperations == 0 && _queue.front() == _taskId` — it is the oldest waiter and nothing else is active, so it pops the queue head and `++_activeOperations`.

The crucial property: **pre-calc runs while holding the queue slot**. Because `EnterOperation` is called before pre-calc and before any `IFileSystem::*` call, queued tasks never interleave their pre-calc + operation phases. `LeaveOperation` decrements `_activeOperations` and calls `NotifyQueueChanged` to wake the next waiter; it is invoked both on the normal completion path and on the cancel-before-execute fast exit in `ThreadMain`.

`ShouldQueueNewTask` decides the initial `_waitForOthers` for a brand-new task: it returns true when global queue mode (`_queueNewTasks`, true by default) is on **and** `HasActiveOperations()` is true (any task object exists, or `_activeOperations > 0` / `_queue` non-empty). Treating any existing task as "active" makes rapid multi-task starts race-free into Wait mode.

## PerItemTaskScheduler: the shared bounded worker pool

When a task uses per-item execution with `maxConcurrency > 1`, the individual item transfers are dispatched through a single process-wide `PerItemTaskScheduler` obtained from `GetPerItemTaskScheduler()` (a function-local static). This is shared across *all* active tasks so the total worker count stays bounded and workers can be reassigned between tasks after each item.

- `ensureWorkers()` lazily spawns `std::clamp(hardware_concurrency, 1, kMaxInFlightFiles=16)` `std::jthread` workers on first use.
- A producer calls `StartJob(task, maxConcurrency, totalItems, processIndex)`. The `Job` carries a `nextIndex`/`inFlight` cursor (scheduler-mutex protected) and a `processIndex(size_t)` callback. The producer then blocks in `WaitJob` on the job's `doneCv`.
- Workers loop in `workerMain`: wait on `_cv` until `hasSchedulableWorkLocked()`, then `tryDequeueWorkLocked` picks a job round-robin (`_rrCursor`), bumps `nextIndex`/`inFlight`, runs `processIndex(index)`, then decrements `inFlight` and re-runs cleanup.
- **Fairness/starvation guards** live in `effectiveMaxConcurrencyLocked`: with one active job and `totalItems > 1` it reserves one worker (`workerCount - 1`) so a second job can start without waiting; with multiple active jobs it caps each job so no single job occupies every worker (`workerCount - activeJobs + 1`). A paused or cancelled task's job is treated as not schedulable, so its workers free up immediately.
- If no workers could be created, `StartJob` degrades to running every `processIndex` inline on the calling thread.

`Shutdown` requests stop on all workers, joins them **before** clearing the job list (so stack-backed operation contexts cannot unwind while a callback is still active), then finishes any remaining jobs. The state cancels producers before shutting the scheduler down; `RunFileOpsPerItemSchedulerShutdownQuietPointSelfTestForSelfTest` proves the join-before-unwind ordering.

## Wait vs Parallel vs queue-pause

Three distinct mechanisms cooperate; do not conflate them:

| Mechanism | Owner state | Effect | Where it blocks |
|-----------|-------------|--------|-----------------|
| **Start gating** (Wait mode) | `Task::_waitForOthers`, `_queue`, `_activeOperations` | A waiting task blocks *before* pre-calc until it is the oldest waiter and nothing is active | `EnterOperation` on `_queueCv` |
| **Pause** (user) | `Task::_paused` | Worker stalls at the next progress checkpoint | `WaitWhilePaused` / `WaitWhilePreCalcPaused` on `_pauseCv` |
| **Queue pause** (mode-driven) | `Task::_queuePaused` | Already-active non-oldest tasks stall when switching to Wait while many are running | same `WaitWhilePaused` path |

`ApplyQueueMode(queue)` sets `_queueNewTasks` and reconciles existing tasks: in Parallel it clears `_waitForOthers` on everything (unblocking start-gated tasks); in Wait it sets `_waitForOthers` on not-yet-started tasks, then calls `UpdateQueuePausedTasks`.

`UpdateQueuePausedTasks` selects the single task that keeps running in Wait mode using **operation-enter time, not task ID**: it walks all tasks that `HasEnteredOperation()`, picks the smallest `GetEnteredOperationTick()` (ties broken by ID), clears `_queuePaused` on that one, and sets `_queuePaused` on the rest. Tasks that have not yet entered are left unpaused (they are still gated at `EnterOperation`). This gives deterministic "oldest active task continues" serialization.

`WaitWhilePaused` loops while either `_paused` or `_queuePaused` is set, waiting on `_pauseCv`; it also handles conflict convergence (below). `WaitWhilePreCalcPaused` is the pre-calc-phase equivalent and additionally returns when `_preCalcSkipped` is set.

## 5F early admission: the overlap latch

For **Copy** with pre-calc enabled, the engine overlaps the recursive size scan with the byte transfer so first-byte latency on deep trees does not wait for the full scan. In `ThreadMain`:

```cpp
const bool useEarlyAdmission =
    (_operation == FILESYSTEM_COPY && _enablePreCalc && ! _preCalcSkipped.load(...));
```

When `useEarlyAdmission`, `RunPreCalculation` is launched on a **side `std::jthread`** via `TryStartPreCalculationThread` while `ExecuteOperation` runs the transfer concurrently. If the side thread cannot be created, it falls back to the serial pre-calc-then-execute path. **Move and Delete always stay serial** because they mutate/remove the source, and a concurrent scan would size a tree that is being deleted.

Both producers raise the shared totals under `_progressMutex` with `std::max`, so totals/ETA stay "estimating" (gated on `_preCalcCompleted`) until pre-calc publishes, then reconcile in place.

The deterministic proof that the overlap happened is the `_transferStartedBeforePreCalcComplete` latch. It is set inside `FileSystemProgress` (a *transfer* progress callback) the first time a transfer callback fires while pre-calc is still in progress and not yet complete:

```cpp
if (! _transferStartedBeforePreCalcComplete.load(...) &&
    _preCalcInProgress.load(...) && ! _preCalcCompleted.load(...))
{
    _transferStartedBeforePreCalcComplete.store(true, ...);
}
```

This state is impossible under the old serial model, so a self-test can assert it to prove the transfer genuinely overlapped the scan.

## Conflict call sequence

`FileSystemIssue` is the `IFileSystemCallback` method plugins call when an item fails and needs a user decision. The engine arbitrates through `Task::_conflictArbiter` (a `ConflictArbiter`: a mutex, a `cv`, a per-bucket `decisionCache`, a single `ConflictPromptState prompt`, an `ownerThreadId`, and a `wil::unique_event` `decisionEvent`). The full sequence:

1. `FileSystemIssue` calls `WaitWhilePaused`, then classifies the failure into a `ConflictBucket` (`ClassifyConflictBucket` — Exists, ReadOnly, AccessDenied, SharingViolation, DiskFull, PathTooLong, RecycleBinFailed, NetworkOffline, UnsupportedReparse, Unknown).
2. It first consults the per-task per-bucket cache via `LoadConflictDecisionFromCache`. A cached decision (set only by "Apply to all" or "Skip All") short-circuits without prompting.
3. Otherwise it computes retry eligibility from the `PerItemCallbackCookie::issueRetryCounts` (Retry is capped to one per `(item, bucket)`), then calls **`BeginConflictPrompt`**. Under the arbiter mutex this re-checks the cache, waits until no other prompt is active (`prompt.active == false`) so **at most one prompt per task** is live, then publishes the `ConflictPromptState` (bucket, paths resolved to the most-specific child, the action layout, `ownerThreadId = GetCurrentThreadId()`) and resets `decisionEvent`. It returns `{action, ownsPrompt}`.
4. If `ownsPrompt`, the worker calls **`WaitForConflictDecision`**, which loops `WaitForSingleObject(decisionEvent, 50)` (polling cancellation between waits) until the popup signals it.
5. The popup, on a button click, calls **`Task::SubmitConflictDecision(action, applyToAllChecked)`** (routed from `FolderWindow.FileOperations.Popup.cpp` after `FindTask`). It writes `decisionAction` / `decisionApplyToAll` under the arbiter mutex (forcing `applyToAll = false` for Retry) and `SetEvent(decisionEvent)`.
6. `WaitForConflictDecision` reads the decision, caches it when `applyToAll` and the action is cacheable, and calls `ClearConflictPrompt` (which clears `prompt`, resets `ownerThreadId`, and notifies the arbiter cv).
7. `FileSystemIssue` maps the resulting `ConflictAction` to a `FileSystemIssueAction` out-param (Overwrite / ReplaceReadOnly / PermanentDelete / Retry / Skip / Cancel) and returns. `Skip`/`SkipAll` set `_observedSkipAction`; `SkipAll` additionally stores `Skip` in the cache.

While one worker owns the prompt, the **other in-flight workers for the same task converge to a stopped state**: `WaitWhilePaused` checks `prompt.active && ownerThreadId != 0 && ownerThreadId != currentThreadId` and blocks those workers on the arbiter cv until the prompt clears. This satisfies the spec's "serialize prompts; all in-flight workers converge while waiting" rule. Both the serial per-item loop and the cross-FS bridge use the identical `BeginConflictPrompt` / `WaitForConflictDecision` / `StoreConflictDecisionInCache` calls.

### Invariants

- No implicit overwrite/replace-readonly/continue-on-error — conflicts must surface.
- Overwrite and Replace-read-only grants are **one-shot** (single item) unless "Apply to all" is checked.
- Retry is per-`(item, bucket)`, capped to one, and **never cached**.
- Directory-vs-directory existence is a merge, not a conflict; a colliding child raises a conflict named by its leaf.

## Cross-filesystem bridge internals and data safety

When the source and destination panes resolve to different `IFileSystem` instances, `ExecuteOperation` detects `_destinationFileSystem != nullptr` for Copy/Move and drives the host's `CrossFileSystemBridge` (defined inside `ExecuteOperation` in `State.cpp`) instead of delegating to one plugin. The bridge `QueryInterface`s `IFileSystemIO` on both providers (source read, destination write) and `IFileSystemDirectoryOperations` on the destination (directory create). Per-worker concurrency is the min of the source and destination per-item budgets.

Callbacks into the host go through a `BridgeCallback` that holds a `callbackMutex` so `IFileSystemCallback` invocations are serialized; each concurrent transfer uses a distinct `progressStreamId`. `FileSystemProgress` on this callback calls `WaitWhilePaused` and returns `ERROR_CANCELLED` on cancel, preserving pause/cancel responsiveness.

### Staging, promote, rollback

The bridge never writes the final destination path directly. For each file:

1. **Stage** — `MakeTempDestinationPath` builds `<dest>.rs_tmp_<128-bit-CSPRNG>_<streamId>` (128 bits from `BCryptGenRandom`, falling back to PID/TID/tick only if the CSPRNG fails). `CreateFileWriter` opens that temp with overwrite/replace-readonly flags **stripped** (`tempFlags`), so a grant for the final destination can never authorize clobbering a staging path.
2. **Copy** — bytes are pumped through a host buffer (serial `copySerial` or a producer/consumer pipeline mode). A `wil::scope_exit cleanupTemp` deletes the temp via `BestEffortDeleteTempFile` unless `promoted` — this is the **rollback**: a failed or cancelled transfer leaves no temp residue.
3. **Integrity gate** — after the copy, if the source size was known and `fileCompletedBytes != fileTotalBytes`, the file finishes `ERROR_PARTIAL_COPY` (no promote).
4. **Commit + promote** — `writer->Commit()`, then `PromoteTempToFinalPath` does `destinationFs.MoveItem(temp -> final)`, carrying the resolved overwrite/replace-readonly grant into the promote flags so the rename can replace an existing file. Only after a successful promote is `promoted = true` (suppressing the rollback delete).

### Move data safety

Cross-filesystem Move is Copy + Delete, and the source is deleted only after the destination is **byte-proven**:

- If `IFileReader::GetSize` on the source fails for a Move, the bridge returns `ERROR_PARTIAL_COPY` and **preserves the source** rather than guessing.
- After a successful promote, for Move with a known source size, the bridge re-opens the destination (`CreateFileReader` + `GetSize`) and requires `destinationSize == fileTotalBytes`. On mismatch or re-stat failure it returns `ERROR_PARTIAL_COPY` with a "preserving source" diagnostic — the source delete is skipped.
- A child answered **Skip** increments `skippedFileConflictCount`; the source for that child is never deleted and the whole transfer ends `ERROR_PARTIAL_COPY` rather than failing the directory closed.

`ThreadMain` surfaces `ERROR_PARTIAL_COPY` as a Warning diagnostic with explicit "source preserved; partial copy left" wording for Move, so both sides are flagged for user review in the Issues pane. Buffer size is snapshotted from `fileOperations.crossFsBridgeBufferSizeKB` at task creation and may be adapted upward from `GetTransferHints`.

## Testing and instrumentation

Every change to liveness or data safety must add deterministic `--fileops-selftest` coverage (the engine exposes `ENABLE_TESTS` hooks: `SetFileOpsBridgePipelineModeForSelfTest`, `SetFileOpsBridgeFailNextFileCopiesForSelfTest`, `SetFileOpsBridgeFailNextSourceGetSizeForSelfTest`, `SetFileOpsPreCalcThreadStartFailureForSelfTest`, `SetFileOpsAutoConcurrencyOverrideForSelfTest`, `SetFileOpsPostFinishedCompletionPauseForSelfTest`, and the scheduler shutdown quiet-point runner). Perf is emitted through `Debug::Perf` under the `FileOps.*` metric families (`FileOps.Queue.*`, `FileOps.Scheduler.*`, `FileOps.PreCalc`, `FileOps.Operation`, `FileOps.Bridge.*`, `FileOps.Conflict.WaitUs`). The Riptide and Floodgate WIP plans (`Specs/Plans/WIP/Operation_Riptide_FairstreamRemediation_DataSafetyConflictParity_2026-06-15.md`, `Operation_Floodgate_RiptideMergeCloudDataSafetyAndCloseout_2026-06-17.md`) track remediation of cross-FS data-safety and conflict-routing parity.

## See also

- [../FileOperations.md](../FileOperations.md) — user-facing Copy/Move/Delete behavior and the progress popup.
- [../DeveloperGuide.md](../DeveloperGuide.md) — the File Operations Engine and Plugin Host Model overview sections.
- [Specs/FileSystem/FileSystem_FileOperations.md](../../Specs/FileSystem/FileSystem_FileOperations.md) — the normative execution model, conflict, bridge, and data-safety contract.
- [Specs/Plugins/Plugins_VirtualFileSystem.md](../../Specs/Plugins/Plugins_VirtualFileSystem.md) — `IFileSystem` capability and `pathIdentity` contracts.
