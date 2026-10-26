# File Operations Engine (Developer)

The File Operations Engine runs every long-running Copy/Move/Delete off the UI thread, gates them through a Queue/Parallel queue, dispatches per-item work on a shared scheduler, arbitrates conflicts inline, and moves bytes between unrelated providers through a host-driven cross-filesystem bridge. This page documents the internals; for the user-facing behavior see [../FileOperations.md](../FileOperations.md), for the higher-level subsystem overview see [../DeveloperGuide.md](../DeveloperGuide.md), and for the normative contract see [Specs/FileSystem/FileSystem_FileOperations.md](../../Specs/FileSystem/FileSystem_FileOperations.md).

## Source layout

The engine is split across one header and several translation units so a single .cpp does not grow unbounded. All symbols live under `FolderWindow::FileOperationState` and its nested `Task`.

| File | Role |
|------|------|
| [`FolderWindow.FileOperationsInternal.h`](../../RedSalamander/FolderWindow.FileOperationsInternal.h) | Private `FileOperationState`, `Task`, result and admission types. |
| [`FolderWindow.FileOperations.cpp`](../../RedSalamander/FolderWindow.FileOperations.cpp) | UI-thread `AdmitOperation` and `AdmitInlineRename`: captured endpoint facts, planning, capability and same-object checks before `StartOperation`. |
| [`FolderWindow.FileOperations.State.Runtime.cpp`](../../RedSalamander/FolderWindow.FileOperations.State.Runtime.cpp) | `StartOperation`, queue-mode changes, lifecycle, cancellation and popup telemetry-work predicate. |
| [`FolderWindow.FileOperations.State.Queue.cpp`](../../RedSalamander/FolderWindow.FileOperations.State.Queue.cpp) | Queue entry/exit, completion posting, fallback storage and worker-quiet drain. |
| [`State.Conflict.cpp`](../../RedSalamander/FolderWindow.FileOperations.State.Conflict.cpp), [`FileOperationsConflictBroker.cpp`](../../RedSalamander/FileOperationsConflictBroker.cpp) | Generation-checked consent binding and one-shot broker decisions. |
| [`FolderWindow.FileOperations.State.cpp`](../../RedSalamander/FolderWindow.FileOperations.State.cpp) | Task execution, provider/bridge mutation, callbacks, typed-result reduction and conflict waits. |
| [`FolderWindow.FileOperations.State.Diagnostics.cpp`](../../RedSalamander/FolderWindow.FileOperations.State.Diagnostics.cpp) | Completed task summary and diagnostic capture. |
| [`Popup.cpp`](../../RedSalamander/FolderWindow.FileOperations.Popup.cpp), [`Popup.PathDetail.cpp`](../../RedSalamander/FolderWindow.FileOperations.Popup.PathDetail.cpp), [`IssuesPane.cpp`](../../RedSalamander/FolderWindow.FileOperations.IssuesPane.cpp), [`IssuesProjection.cpp`](../../RedSalamander/FolderWindow.FileOperations.IssuesProjection.cpp) | Task/Issues presentation, path-detail focus, lifecycle and timer policy. |

## Per-task jthread model

UI commands first call `AdmitOperation` (or `AdmitInlineRename`) in [`FolderWindow.FileOperations.cpp`](../../RedSalamander/FolderWindow.FileOperations.cpp), which captures provider/endpoint identity, checks capabilities and overlap, and builds typed plans. `StartOperation` in [`State.Runtime.cpp`](../../RedSalamander/FolderWindow.FileOperations.State.Runtime.cpp) accepts that prepared admission, snapshots task settings, publishes the task, and starts its worker. Empty selected input returns without creating a task.

```cpp
rawTask->_thread = std::jthread([rawTask](std::stop_token stopToken) noexcept { rawTask->ThreadMain(stopToken); });
```

There is one `std::jthread` per task (`Task::_thread`). The `std::stop_token` is the cooperative cancellation channel: `ThreadMain` installs a `std::stop_callback` that, on stop request, wakes `_pauseCv`, the `_conflictArbiter.cv`, and the state's `_queueCv` so a task blocked anywhere unblocks promptly. Cancellation itself is requested via `Task::RequestCancel` (sets `_cancelled`); both `_cancelled` and the stop token are polled at every checkpoint.

`ThreadMain` is the control-flow spine:

1. `SetWaitingInQueue(_waitForOthers)` so the UI shows "Waiting" immediately for queued tasks.
2. `_state->EnterOperation(*this, stopToken)` — the start gate (below). If it returns `false` the task was cancelled while queued; it stores `ERROR_CANCELLED` and calls `PostCompleted`.
3. On success it latches `_enteredOperationTick` / `_enteredOperation` (used by queue-pause ordering).
4. Prepares the selected inputs and uses the provider/host shared discovery path to publish growing facts while work is scheduled.
5. `ExecuteOperation()` — the actual transfer/delete.
6. Waits for owned producers/workers to become quiet, records discovery/operation/bridge metrics and typed diagnostics, then publishes terminal task state. `PostCompleted` first tries an owner-window payload; failed posts retain the completion for owner-thread adoption only after the worker is joined.

On `Shutdown` the state requests cancellation and joins owned workers and completion drains before releasing borrowed task state. A normal main-window close with active work first asks for Cancel all, keeps the window alive while cards show Stopping, and resumes `WM_CLOSE` after the tasks are quiet; D10 exercises that application path. While that exit is pending, `StartOperation` and `AdmitScheduledRename` refuse every new admission with `FolderWindow::kFileOperationRefusedForPendingExit` before any prompt (and `StartOperation` again after its last prompt), and a second close attempt only brings the popup forward. The exit stays pending until the resumed `WM_CLOSE` has been handled (`ConfirmCancelAllFileOperations` ends it, or the `WndMsg::kFileOperationsDeferredCloseSettled` message posted right behind the `WM_CLOSE`), so a posted message queued ahead of that close is refused rather than admitted as work the close would ask to cancel again. The check is `_owner.IsFileOperationsExitPending()`, which on the Compare Directories window's own `FolderWindow` follows the application `FolderWindow` (`SetApplicationExitFolderWindow`), so Compare's separate engine refuses too. The same link registers Compare's `FolderWindow` with the application one (`_exitLinkedFolderWindows`): the exit consent counts every linked engine, `CancelAllFileOperationsThenClose` cancels them all, and the deferred close resumes only when all are quiet (a linked engine's `ResumeDeferredCloseIfFileOperationsQuiet` forwards to the application window). The Compare window's own `WM_CLOSE` calls `ConfirmCancelAllFileOperations` on its `FolderWindow`, which asks to cancel that window's work and close it and defers that close the same way (FileOps self-test `FileOps_CompareWindowCloseConsent`).

## EnterOperation: the start-gating queue

`EnterOperation` (in `FolderWindow.FileOperations.State.Queue.cpp`) implements Wait vs Parallel start gating against `FileOperationState`-owned state: `_queueMutex`, `_queueCv`, a FIFO `std::deque<uint64_t> _queue` of waiting task IDs, and an `unsigned long _activeOperations` count.

- If the task's `_waitForOthers` is false (Parallel), it simply `++_activeOperations` and returns `true` — no blocking.
- If `_waitForOthers` is true (Wait), it pushes its task ID onto `_queue`, calls `NotifyQueueChanged`, and waits on `_queueCv` until one of:
  - the stop token fires or `_cancelled` is set (returns `false` after `RemoveFromQueue`), or
  - `_waitForOthers` flipped to false mid-wait (mode switched to Parallel) — it removes itself from the queue and proceeds, or
  - `_activeOperations == 0 && _queue.front() == _taskId` — it is the oldest waiter and nothing else is active, so it pops the queue head and `++_activeOperations`.

The start gate precedes discovery and provider mutation. A queued task cannot begin its traversal or transfer before admission to the active operation set.

`ShouldQueueNewTask` decides the initial `_waitForOthers` for a brand-new task: it returns true when global queue mode (`_queueNewTasks`, true by default) is on **and** `HasActiveOperations()` is true (any task object exists, or `_activeOperations > 0` / `_queue` non-empty). Treating any existing task as "active" makes rapid multi-task starts race-free into Wait mode.

## PerItemTaskScheduler: the shared bounded worker pool

When a task uses per-item execution with `maxConcurrency > 1`, the individual item transfers are dispatched through a single process-wide `PerItemTaskScheduler` obtained from `GetPerItemTaskScheduler()` (a function-local static). This is shared across *all* active tasks so the total worker count stays bounded and workers can be reassigned between tasks after each item.

- `ensureWorkers()` lazily spawns `std::clamp(hardware_concurrency, 1, kMaxInFlightFiles=16)` `std::jthread` workers on first use.
- A producer calls `StartJob(task, maxConcurrency, totalItems, itemPolicy, firstFailure)`. The `Job` carries a `nextIndex`/`inFlight` cursor (scheduler-mutex protected) plus non-owning pointers to the stack-bound policy and failure slot. The producer then blocks in `WaitJob` on the job's `doneCv`; `WaitJob` and `Shutdown` are the quiet-point guarantee that those references cannot outlive their owners.
- Workers loop in `workerMain`: wait on `_cv` until `hasSchedulableWorkLocked()`, then `tryDequeueWorkLocked` picks a job round-robin (`_rrCursor`), bumps `nextIndex`/`inFlight`, runs `itemPolicy.Process(index)`, records the first failure/cancellation through the scheduler, then decrements `inFlight` and re-runs cleanup.
- **Fairness/starvation guards** live in `effectiveMaxConcurrencyLocked`: with one active job and `totalItems > 1` it reserves one worker (`workerCount - 1`) so a second job can start without waiting; with multiple active jobs it caps each job so no single job occupies every worker (`workerCount - activeJobs + 1`). A paused or cancelled task's job is treated as not schedulable, so its workers free up immediately.
- If no workers could be created, `StartJob` degrades to running every `itemPolicy.Process(index)` inline on the calling thread.

`Shutdown` requests stop on all workers, joins them **before** clearing the job list (so stack-backed operation contexts cannot unwind while a callback is still active), then finishes any remaining jobs. The state cancels producers before shutting the scheduler down; `RunFileOpsPerItemSchedulerShutdownQuietPointSelfTestForSelfTest` proves the join-before-unwind ordering.

## Wait vs Parallel vs queue-pause

Three distinct mechanisms cooperate; do not conflate them:

| Mechanism | Owner state | Effect | Where it blocks |
|-----------|-------------|--------|-----------------|
| **Start gating** (Wait mode) | `Task::_waitForOthers`, `_queue`, `_activeOperations` | A waiting task blocks *before* discovery/mutation until it is the oldest waiter and nothing is active | `EnterOperation` on `_queueCv` |
| **Pause** (user) | `Task::_paused` | Worker stalls at the next cooperative progress checkpoint | `WaitWhilePaused` on `_pauseCv` |
| **Queue pause** (mode-driven) | `Task::_queuePaused` | Already-active non-oldest tasks stall when switching to Wait while many are running | same `WaitWhilePaused` path |

`ApplyQueueMode(queue)` sets `_queueNewTasks` and reconciles existing tasks: in Parallel it clears `_waitForOthers` on everything (unblocking start-gated tasks); in Wait it sets `_waitForOthers` on not-yet-started tasks, then calls `UpdateQueuePausedTasks`.

`UpdateQueuePausedTasks` selects the single task that keeps running in Wait mode using **operation-enter time, not task ID**: it walks all tasks that `HasEnteredOperation()`, picks the smallest `GetEnteredOperationTick()` (ties broken by ID), clears `_queuePaused` on that one, and sets `_queuePaused` on the rest. Tasks that have not yet entered are left unpaused (they are still gated at `EnterOperation`). This gives deterministic "oldest active task continues" serialization.

`WaitWhilePaused` loops while either `_paused` or `_queuePaused` is set, waiting on `_pauseCv`; it also handles conflict convergence (below).

## Streaming discovery and scheduling

Copy, Move and Delete share traversal work with counting instead of running an
independent full-tree pre-calculation pass. The local FileSystem provider publishes
discovery facts as its producer walks, validates and schedules work. Host bridge
routes preserve the same open/closed discovery distinction. Local selected files
receive fresh metadata checks before payload transfer; pane-displayed sizes are
not admission authority. Read failures remain per-item facts.

A provider listing closes a folder's discovery once all its entries and recursive
producers are accounted for. Folder operation completion is separate: descendants
and metadata restoration may still be running. A growing count is provisional;
percentage/ETA and footer/taskbar state must follow the explicit closure facts.
The provider's lifetime/join rules remain part of mutation safety, especially
when Move removes source leaves or a Delete retains parent handles.

**Skip counting** is a one-way scheduling change. It stops discovery running ahead
and releases its reserved transfer capacity, while required enumeration, name
validation and conflict checks continue. It never discards items or fabricates a
closed total. The old side pre-calculation thread and overlap latch are not the
current architecture.

Use the authoritative
[streaming discovery and scheduling contract](../../Specs/FileSystem/FileSystem_FileOperations.md#streaming-discovery-and-scheduling)
for queue bounds, lifetime, cancellation and route-specific behavior. The popup
shows completed/found counts, current bytes and the effective automatic transfer
budget from snapshots; it does not derive completion from selected-root counts.
Speed is a rolling wall-clock observation. Terminal Last note text is measured
and wrapped above progress/actions, including after resize.

## Conflict call sequence

A provider issue enters the task's conflict arbiter. The worker classifies the issue and reuses an eligible per-task decision cache; otherwise BeginConflictPrompt publishes one task prompt with a new generation and waits for an actual decision. Other workers for that task converge at the pause/conflict barrier.

Every hosted action, overflow action, Escape path and UIA activation submits its captured prompt generation and action placement. `FileOperationsConflictBroker::TryAcceptLocked` checks that generation, the currently offered action and one-shot state under the arbiter mutex, and signals the event before unlocking. `WaitForConflictDecision` ignores an event without accepted state, then consumes the decision, updates any allowed repeat policy, resets the prompt/event and releases ownership under the same lock. Cancellation is a separate durable request; a stale signal cannot invent Cancel. Held Enter/Escape activation repeats are filtered before they reach a newly published prompt; navigation autorepeat remains available. See [`FileOperationsConflictBroker.cpp`](../../RedSalamander/FileOperationsConflictBroker.cpp), [`State.Conflict.cpp`](../../RedSalamander/FolderWindow.FileOperations.State.Conflict.cpp), [`State.cpp`](../../RedSalamander/FolderWindow.FileOperations.State.cpp) and [`Popup.cpp`](../../RedSalamander/FolderWindow.FileOperations.Popup.cpp).

### Invariants

- No implicit overwrite/replace-readonly/continue-on-error — conflicts must surface.
- Overwrite and Replace-read-only grants are **one-shot** (single item) unless "Apply to all" is checked.
- Retry is per-`(item, bucket)`, capped to one, and **never cached**.
- Applying a cached "Apply to all" choice records nothing per item: every application of one originating prompt folds into one Issues record (`appliedCount`, last item, the items' common typed result, the first application's time), refreshed at most every 250 ms and at the task's next user prompt, and written to the JSONL log once when the task completes. The bounded Issues reservoirs evict notes, then records whose row has no retained problem, then whole problem rows oldest first: a record never evicts a real Warning/Error, and a retained prompt keeps its decision and result.
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
- After successful promote, Move compares the source-byte BLAKE3 digest with a bound provider BLAKE3 proof or host readback of the published destination before removing that leaf's original source. Matching size alone is insufficient. If proof is unavailable or mismatched, that source is retained; other leaves keep their independent results.
- A child answered **Skip** increments `skippedFileConflictCount`; the source for that child is never deleted and the whole transfer ends `ERROR_PARTIAL_COPY` rather than failing the directory closed.

`ThreadMain` reports `ERROR_PARTIAL_COPY` with item-level typed outcomes and diagnostics. A verified, published Managed Move leaf may have its source removed while later failed leaves retain theirs; inspect each outcome before retry or cleanup. Buffer size is snapshotted from `fileOperations.crossFsBridgeBufferSizeKB` at task creation and may be adapted upward from `GetTransferHints`.

## I27 follow-up ownership seams

The [I27 plan](../../Specs/Plans/Done/UI_FileOperationsSafetyRemediation_2026-10-01.md)'s [original gallery and protected interaction](../../Specs/Plans/Done/UI_FileOperationsSafetyRemediation_2026-10-01/implementation/2026-10-04-final-gallery/20261003T232201Z-75488-0a7264eaad564dcd8aa1014d5093c463/gallery.html)
retain frozen build/source bindings, native display metadata and original assertions.
They cover deterministic synthetic File Operations and Issues presentation; protected
foreground interaction has a separate receipt. Later production viewer-menu ownership
changes in `Common/ViewerMenuHost.h` and the PE, Web, Text, Space and ImgRaw adapters,
plus subsequent test and diagnostic fixes, require their own source-bound native
receipts. These original captures do not qualify those later inputs or final Fresh Full.

- [`FileOperationConfirmation.cpp`](../../RedSalamander/FileOperationConfirmation.cpp) formats Copy/Move and permanent-Delete locations using the captured provider path profile. Unknown or mixed domains keep raw spelling or use the bounded Multiple locations summary; formatting does not change execution endpoints.
- The host's typed result keeps completion, destination publication, verification, source disposition and owned-stage cleanup independent. [`FileOperationItemResultReducer.h`](../../RedSalamander/FileOperationItemResultReducer.h) classifies provider receipts; host-managed proofs retain their own authority. A task-level partial failure does not prove every source was retained or every destination was published. A claimed successful Delete receipt that commits but retains the original is a failed contradiction. A canceled or partial recursive Delete can legitimately retain its root. Per-item execution marks a selected root `MutationPossible` only when `QualifiedItemPolicy::Process` dispatches it (`Task::MarkSourceItemMutationPossible`); `ThreadMain`'s `EnterSourceItemsMutationGate` keeps the task-wide mark only for bulk and rename executors, so roots a Cancel or a stopping failure left undispatched finalize `NotAttempted`/`Retained`/`Canceled`, not Unknown and not with the failing root's status (`FinalizeTypedItemResults` reduces them with `ERROR_CANCELLED`).
- `PostCompleted` in [`State.Queue.cpp`](../../RedSalamander/FolderWindow.FileOperations.State.Queue.cpp) stores failed-post completions until owner-thread adoption after worker quiet, including when the reaper or its wakeup cannot be posted.
- [`Popup.cpp`](../../RedSalamander/FolderWindow.FileOperations.Popup.cpp) suspends periodic telemetry when no operation or visible animation needs it, publishes the final taskbar state once and retains active hidden/minimized taskbar updates. [`AdmissionBaseline.cpp`](../../RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.Tests.AdmissionBaseline.cpp) measures actual UI-thread admission. The retained [F6/v2](../../Specs/TestRuns/4cb089111a23/FileOps/20261003T211512Z-80796-0cb73e59c3c741f9bcae95a38bb43929/review/ROOT-F6-V2-BOUNDED-ASSESSMENT.md) and [F7](../../Specs/TestRuns/4cb089111a23/FileOps/20261003T191317Z-82816-199e9f25af4d49739d1410c6881e3e0d/review/ROOT-F7-ASSESSMENT.md) assessments accept bounded measured behavior; real mounted routes remain capability-skipped and unqualified after access denied. Local's new existing-path observation reuses its brief route cache only after fresh endpoint and mount GUID agreement; unavailable observations keep the original volume resolver. Its native correctness and paired timing qualification remain open, tracked in [I29 Q2-L](../../Specs/Plans/WIP/FileOperations_PostMergeQualification_2026-10-05.md). Full, protected/gallery and archive governance retain separate gates.
- The private D10 child in [`SelfTest.OwnedShutdown.cpp`](../../RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.Tests.OwnedShutdown.cpp) uses normal Copy overwrite consent, a positive incomplete host bridge stage write, actual `WM_CLOSE`/cancellation and natural application teardown. Its parent checks the external tree only after child exit. The production defer/resume path is in [`FolderWindow.FileOperations.cpp`](../../RedSalamander/FolderWindow.FileOperations.cpp). This fixture is evidence only when its retained runtime receipt passes.

## Testing and instrumentation

Every change to liveness or data safety must add deterministic `--fileops-selftest` coverage (the engine exposes `ENABLE_TESTS` hooks: `SetFileOpsBridgePipelineModeForSelfTest`, `SetFileOpsBridgeFailNextFileCopiesForSelfTest`, `SetFileOpsBridgeFailNextSourceGetSizeForSelfTest`, `SetFileOpsPreCalcThreadStartFailureForSelfTest`, `SetFileOpsAutoConcurrencyOverrideForSelfTest`, `SetFileOpsPostFinishedCompletionPauseForSelfTest`, and the scheduler shutdown quiet-point runner). Perf is emitted through `Debug::Perf` under the `FileOps.*` metric families (`FileOps.Queue.*`, `FileOps.Scheduler.*`, `FileOps.PreCalc`, `FileOps.Operation`, `FileOps.Bridge.*`, `FileOps.Conflict.WaitUs`). The File Operations safety remediation record and retained TestRuns receipts document qualification; [plan lifecycle](../../Specs/README.md#plan-lifecycle) governs its WIP/Done location. Normative execution, popup, self-test and performance requirements stay in their domain contracts; the old Riptide/Floodgate WIP paths are retired history.

## See also

- [../FileOperations.md](../FileOperations.md) — user-facing Copy/Move/Delete behavior and the progress popup.
- [../DeveloperGuide.md](../DeveloperGuide.md) — the File Operations Engine and Plugin Host Model overview sections.
- [Specs/FileSystem/FileSystem_FileOperations.md](../../Specs/FileSystem/FileSystem_FileOperations.md) — the normative execution model, conflict, bridge, and data-safety contract.
- [Specs/Plugins/Plugins_VirtualFileSystem.md](../../Specs/Plugins/Plugins_VirtualFileSystem.md) — `IFileSystem` capability and `pathIdentity` contracts.
