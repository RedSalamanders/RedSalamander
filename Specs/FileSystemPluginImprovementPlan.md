# Filesystem Plugin Improvement Plan (Resiliency + Scalability + Parallel Ops)

Last updated: 2026-02-03

## Goal

Improve the Filesystem plugin’s:

- **Resiliency**: fewer missed directory change events under load; make event loss **detectable** and drive a **host resync**.
- **Scalability**: handle very large directory enumerations without hard failure (and without retaining huge buffers unnecessarily).
- **Performance**: parallelize batch operations safely with correct **progress semantics** and **cancellation**.

Scope includes the built-in local filesystem plugin (`Plugins/FileSystem`) and small host-side behavior in `RedSalamander` to perform a refresh when change notifications overflow.

## Non-goals (v1)

- Rewriting the host directory model / cache (only plug into existing refresh mechanisms).
- Replacing file copy/move primitives (stick to Win32 APIs already used).
- Making directory watch “lossless” (Windows change notification is best-effort; overflows are inevitable).

## Key Constraints (from AGENTS.md + skills)

- **RAII mandatory** for Windows resources: use WIL wrappers (`.github/skills/wil-raii/SKILL.md`).
- **Callbacks are raw vtables** (not COM): cookie passthrough; plugin holds weak pointers (`.github/skills/plugin-callbacks/SKILL.md`).
- **Do not block UI thread**; marshal across threads using message patterns (`.github/skills/async-threading/SKILL.md`, `.github/skills/win32-wndproc/SKILL.md`).
- **Error handling**: Win32 failures use `Debug::ErrorWithLastError(...)` + `HRESULT_FROM_WIN32` (`.github/skills/error-handling/SKILL.md`).
- **No global /wd suppressions**; fix warnings locally (`.github/skills/compiler-warnings/SKILL.md`).
- **C++23 style**: no raw `new/delete`, no `goto`, prefer clear STL patterns (`.github/skills/cpp-modern-style/SKILL.md`).

## Current State (as of 2026-02-03)

### Directory watch

File: `Plugins/FileSystem/FileSystem.Watch.cpp`

- Uses a small buffer pool and re-issues `ReadDirectoryChangesW` **before** invoking the host callback (I/O completion is decoupled from parsing/callback delivery).
- On `ERROR_NOTIFY_ENUM_DIR` it sets `notification.overflow = TRUE` (already exists in the interface).
- Hard-caps changes delivered per callback (`kMaxChanges = 128`); exceeding the cap sets `overflow=TRUE`.

### Directory enumeration

File: `Plugins/FileSystem/FileSystem.DirectoryOps.cpp`

- Uses `NtQueryDirectoryFile` and stores results in an in-memory contiguous `FileInfo` buffer.
- Uses progressive growth up to a soft cap and resumes enumeration on growth (no restart); falls back to a higher hard cap before returning `ERROR_INSUFFICIENT_BUFFER`.
  - defaults: 512 MiB soft / 2048 MiB hard
  - configurable via plugin configuration:
    - `enumerationSoftMaxBufferMiB` (default 512)
    - `enumerationHardMaxBufferMiB` (default 2048)
- Calls `MaybeTrimBuffer()` after `ReadDirectoryInfo()` completes to drop enumeration scratch state and trim large allocations.
- `ComputeEntrySize`/`AlignUp` currently aligns to 4 bytes (`sizeof(unsigned long)`).
- `FileInfo` layout is required to match `FILE_FULL_DIR_INFO` (static_assert present).

### Batch file operations

File: `Plugins/FileSystem/FileSystem.FileOps.cpp`

- `CopyItems`/`MoveItems` run with bounded internal parallelism across top-level input items:
  - default max concurrency: 4
  - configurable via `copyMoveMaxConcurrency` (max 8)
- `DeleteItems` parallelizes delete with ordering safety across overlapping inputs (children before parents) and supports Recycle Bin deletes with bounded concurrency:
  - default max concurrency: 8
  - default max concurrency (Recycle Bin): 2
  - configurable via `deleteMaxConcurrency` / `deleteRecycleBinMaxConcurrency`
- Callbacks are invoked directly from worker progress checkpoints, but are **serialized** via a per-operation mutex (host must never receive concurrent callback calls).
- Progress struct already supports items + bytes:
  - `IFileSystemCallback::FileSystemProgress(operationType, totalItems, completedItems, totalBytes, completedBytes, ..., currentItemTotalBytes, currentItemCompletedBytes, ..., progressStreamId, ...)`
  - `totalBytes` is typically `0` from the plugin; host pre-calc provides totals when available.
  - `completedBytes` is populated (Copy/Move from progress callbacks; Delete best-effort via file sizes).

## Contract Decisions (write first; code follows the contract)

### A) Directory watch is best-effort; loss must be detectable

Interface: `Common/PlugInterfaces/FileSystem.h`

- Change notification delivery is **best-effort**.
- **If any events are dropped/coalesced** (OS overflow, internal cap, parse failure, queue pressure), the plugin will set `FileSystemDirectoryChangeNotification::overflow = TRUE`.
- **Host behavior requirement**: `overflow==TRUE` means “incremental events are not trustworthy; perform a full resync”.
  - No new callback is required because `overflow` already exists.
  - We should strengthen comments to make “resync required” explicit.

### B) Callback threading: prefer serialized callbacks (plugin-managed)

Reason: avoid silently requiring host thread-safety and avoid re-entrancy hazards.

- **Goal**: the Filesystem plugin invokes `IFileSystemCallback` and `IFileSystemDirectoryWatchCallback` **serially** (never concurrently) for a given operation/watch.
- Background workers MAY call host callbacks, but the plugin MUST serialize them (e.g., per-operation mutex) so the host does not need to be thread-safe.

If the host later wants concurrency, add an explicit capability/opt-in (not part of this plan).

### C) Progress semantics (items + bytes)

#### Copy/Move

- `totalItems`: fixed input count.
- `completedItems`: number of top-level input items completed (**monotonic**, `0..totalItems`).
- `totalBytes`/`completedBytes`:
  - **monotonic** and represent byte progress for the overall operation.
  - For directories, `totalBytes` may be **discovered incrementally** (e.g., as files are enumerated), so it may increase over time.
  - `completedBytes` never decreases and never exceeds `totalBytes` (when `totalBytes` is known/updated).
- `currentSourcePath/currentDestinationPath`:
  - with parallelism, these represent the **current item for a specific progress stream** (e.g., a worker).
  - `progressStreamId` MUST be stable per worker and distinct among concurrently active workers so the host can display stable per-stream lines.
- `currentItemTotalBytes/currentItemCompletedBytes`:
  - reflect the representative item’s file copy progress (when available).

#### Delete

- `totalItems` / `completedItems` are meaningful and monotonic.
- `totalBytes` is typically `0` from the plugin (host pre-calc provides totals when available).
- `completedBytes` is best-effort and monotonic for local filesystem deletes:
  - increments by deleted file sizes when they can be queried
  - directories contribute `0`
- `currentItemTotalBytes/currentItemCompletedBytes` remain `0` for delete (items-focused UI).

### D) “At least once” and ordering guarantees

- File ops:
  - `FileSystemItemCompleted(operationType, itemIndex, ...)` is called **at most once** per item index that started.
  - Completion order may be **out of input order**; `itemIndex` is the stable identity.
- Directory watch:
  - No “at least once” guarantee. Overflows are inevitable; detect via `overflow==TRUE` and resync.

## Confirming `FileInfo` is in-memory only (ABI safety)

Why: alignment/size changes are dangerous if `FileInfo` is serialized, persisted, or depended on across a stable ABI boundary.

Evidence already in code:

- `Common/PlugInterfaces/FileSystem.h` explicitly documents `IFilesInformation` as an in-memory contiguous buffer.
- `Plugins/FileSystem/FileSystem.DirectoryOps.cpp` has:
  - `static_assert(sizeof(FileInfo) == sizeof(FILE_FULL_DIR_INFO))`
  - `static_assert(offsetof(FileInfo, FileName) == offsetof(FILE_FULL_DIR_INFO, FileName))`

Audit steps (do before changing layout assumptions):

- Search for any serialization/persistence of `FileInfo`:
  - `rg -n "FileInfo\\b|GetBuffer\\("`
  - Verify all consumers treat it as an in-memory buffer and do not store it beyond the `IFilesInformation` lifetime.
- Confirm host never writes `FileInfo` to disk/IPC and does not expose it to plugins as a long-lived ABI.
- **Rule**: Do not change `struct FileInfo` layout. Any alignment work must be limited to internal size computations and buffer management.

## Implementation Plan (phased)

### Phase 1 — Spec/Contract updates (documentation-first)

Files:

- `Common/PlugInterfaces/FileSystem.h`
- (Optional) `Specs/FileOperationsSpec.md`, `Specs/FolderWindowSpec.md` (add brief references)

Tasks:

- [x] Update `IFileSystemCallback` comment block to clarify callbacks are invoked **serially** and may block (host-driven Pause/Skip/Cancel).
- [x] Update directory watch docs:
  - [x] `FileSystemDirectoryChangeNotification::overflow` explicitly means “**resync required**”.
  - [x] Clarify watcher event list is best-effort and may be coalesced/dropped.
- [x] Document progress semantics (Copy/Move bytes; Delete items-only; totalBytes may be unknown).
- [x] Add an “ordering” note: item completion order may be out-of-order; `itemIndex` is authoritative.

Acceptance:

- No code changes required yet; but the contract is unambiguous and aligns with what we will implement.

### Phase 2 — Directory watch double-buffering + decoupled processing

File: `Plugins/FileSystem/FileSystem.Watch.cpp`

Design:

- Maintain 1 outstanding `ReadDirectoryChangesW` per watch handle.
- Use a **buffer pool** (2–4 buffers) so we can re-arm immediately on completion.
- Split work into two stages:
  1) I/O completion: swap buffers and re-issue the next read **immediately**.
  2) Processing: parse the completed buffer and invoke the host callback (serialized).

Queue/backpressure:

- Maintain a bounded queue of completed buffers to process.
- On queue overflow or parse anomalies, set `overflow=TRUE` and coalesce until processed.
- Keep callback invocation **serialized** (no concurrent calls to `IFileSystemDirectoryWatchCallback` for a given watch).

Stop/unwatch:

- Ensure `UnwatchDirectory` guarantees no callbacks after it returns (already documented).
- Cancel outstanding I/O with `CancelIoEx` and wait via `WaitForThreadpoolIoCallbacks`.

Tasks:

- [x] Add buffer pool structure and a small processing queue.
- [x] Re-arm reads before invoking host callback.
- [x] Ensure `overflow=TRUE` on:
  - OS overflow (`ERROR_NOTIFY_ENUM_DIR`)
  - internal cap overflow (`kMaxChanges`)
  - queue overflow / parse bounds failures
- [x] Keep all Windows resources under WIL RAII.

Acceptance:

- Under slow host processing, watcher continues re-arming reads and does not stall.
- Overflow is surfaced via `overflow=TRUE` reliably.

### Phase 3 — Host resync behavior on watch overflow

File: `RedSalamander/FolderWindow.cpp` (or wherever directory watch notifications are handled)

Design:

- On `overflow==TRUE`, schedule a full directory refresh (`ReadDirectoryInfo`) for that watched folder.
- Coalesce refresh requests (rate-limit per folder).
- Ensure refresh work is not on UI thread; marshal updates via message pattern.

Tasks:

- [x] Locate directory watch callback handler in the host (`RedSalamander/FolderWatcher.*`).
- [x] Implement “overflow -> resync” path (host already refreshes on any event; overflow is now logged for visibility).
- [x] Add rate limiting/coalescing to avoid refresh storms (per-folder coalescing via `DirectoryInfoCache::notifyPosted`).

Acceptance:

- Large churn produces overflow, and the host reliably refreshes state without UI stalls.

### Phase 4 — Directory enumeration scalability (buffer cap + trimming + alignment validation)

File: `Plugins/FileSystem/FileSystem.DirectoryOps.cpp`

Design:

- Use progressive growth up to a 512MB cap and resume enumeration when growing (no restart) to avoid O(N) re-enumeration passes.
- If the soft cap is exceeded, grow further up to a hard cap as a fallback for extreme cases; if the hard cap is exceeded, return `ERROR_INSUFFICIENT_BUFFER` (or add an additional fallback strategy).
- Aggressively trim large buffers after use (avoid holding hundreds of MB per instance).
- Alignment work:
  - Verify actual alignment requirements for `FILE_FULL_DIR_INFO` chaining and adjust internal `AlignUp` only if validated.
  - Add debug-only assertions (e.g., `NextEntryOffset` sanity and alignment checks).

Tasks:

- [x] Implement resume-on-grow enumeration up to a 512MB soft cap with a higher hard-cap fallback (single pass; no restart).
- [x] Ensure `MaybeTrimBuffer` releases large capacity after enumeration.
- [x] Add a fallback growth path when the soft cap is exceeded (current hard cap: 2GB).
- [x] Validate alignment before changing it; only adjust internal size computations (never struct layout).
  - Added runtime alignment validation when consuming `FILE_FULL_DIR_INFO` chaining (`NextEntryOffset` + buffer offsets must stay aligned).

Acceptance:

- Large directories no longer fail with `ERROR_INSUFFICIENT_BUFFER` in common cases; worst-case falls back rather than failing.
- Memory returns to baseline after operation (no long-lived huge allocations).

### Phase 5 — Parallel batch operations (Copy/Move small; Delete larger)

File: `Plugins/FileSystem/FileSystem.FileOps.cpp`

Design principles:

- Bounded concurrency (prefer a small internal threadpool/queue over `std::execution::par` for better cancellation control).
- Workers MAY invoke host callbacks, but the plugin MUST serialize callback delivery per operation (the host is not required to be thread-safe).
- Host pause semantics rely on blocking inside progress callbacks; therefore work must reach progress checkpoints frequently enough that pausing blocks workers.

Parallelism policy:

- Copy/Move: small parallelism (default 2–4).
- Delete: higher parallelism (default 8), with lower concurrency for Recycle Bin deletes (default 2), and ordering safety:
  - Avoid deleting a parent directory concurrently with children (delete deepest-first / enforce parent-after-children scheduling).

Progress aggregation:

- Use atomics for counters + a small event queue for per-item completion.
- Throttle progress callback rate (per worker), while keeping pause/cancel responsive (e.g., ~20 Hz for Copy/Move and ~10 Hz for Delete).
- Cancellation:
  - Progress checkpoints call `FileSystemShouldCancel` (serialized) and flip a shared atomic stop flag when cancellation is requested.
  - If `continueOnError == false`, the first non-cancel error requests stop for remaining work.

Copy-specific note:

- `CopyFileExW` progress callback is invoked on the copying thread; it must remain fast.
- It should only do minimal bookkeeping and invoke the host callback in a serialized manner (so the host can pause/cancel by blocking).

Tasks:

- [x] Refactor `OperationContext` so mutable shared state is thread-safe (or split into immutable per-item + shared atomics).
- [x] Implement bounded worker queue for CopyItems/MoveItems.
- [x] Implement ordered parallel DeleteItems across input items:
  - schedule parent/child deletes using a dependency graph (children before parents)
  - support Recycle Bin deletes with bounded concurrency
- [x] Serialize callback delivery per operation (mutex) and propagate dynamic bandwidth limit updates.
- [x] Throttle FileSystemProgress callback frequency (per worker; additional global throttling for parallel delete).

Acceptance:

- Copy/Move speed improves on multi-item workloads.
- Progress is monotonic (items, bytes) and cancellation is responsive.
- Delete can run with higher concurrency without directory-structure races.

### Phase 5b — Configuration knobs (implemented)

Files:

- `Plugins/FileSystem/FileSystem.h`
- `Plugins/FileSystem/FileSystem.cpp`

Design:

- Expose safe **concurrency caps** and **enumeration buffer caps** via plugin configuration JSON so power users can tune behavior without code changes.
- Clamp values to conservative maxima to avoid thread explosions and to keep directory enumeration buffers below Win32 API limits.

Keys (defaults):

- `copyMoveMaxConcurrency` (default 4, max 8)
- `deleteMaxConcurrency` (default 8, max 64)
- `deleteRecycleBinMaxConcurrency` (default 2, max 16)
- `enumerationSoftMaxBufferMiB` (default 512)
- `enumerationHardMaxBufferMiB` (default 2048; clamped to >= soft cap and <= 4095 MiB)

Tasks:

- [x] Add configuration schema + parsing for caps.
- [x] Wire caps into Copy/Move/Delete concurrency and enumeration growth logic.

### Phase 6 — Verification + instrumentation

Build:

- `.\build.ps1 -ProjectName FileSystem`
- `.\build.ps1` (full solution)

Stress scenarios:

- Watcher churn:
  - Create/rename/delete thousands of files rapidly under a watched folder.
  - Ensure no crash/hang; overflow triggers resync; steady-state continues.
- Large directory listing:
  - Enumerate directories with many entries and long names.
  - Verify no `ERROR_INSUFFICIENT_BUFFER` hard-fails (or fallback works).
  - Knob coverage: lower `enumerationSoftMaxBufferMiB` / `enumerationHardMaxBufferMiB` to force fallback / hard-cap paths; confirm memory returns to baseline after enumeration (`MaybeTrimBuffer`).
- Parallel ops:
  - Copy/Move: many files (small + large) to see throughput and progress behavior.
  - Delete: deep trees; verify safe ordering and correctness.
  - Knob coverage: vary `copyMoveMaxConcurrency`, `deleteMaxConcurrency`, and `deleteRecycleBinMaxConcurrency` and confirm correctness + monotonic progress.

Instrumentation (optional but recommended):

- [x] Add debug-only counters/timers for:
  - watcher callback latency / queue depth / overflow count (`FileSystem.Watch`)
  - copy/move throughput and cancel latency (`FileOps.Operation`, `FileOps.CancelLatency`)
  - directory enumeration growth / peak buffer size / trim events (`FileSystem.DirectoryOps.Enumerate`, `FileSystem.DirectoryOps.TrimBuffer`)

## Next-session “start here” checklist

- [x] Re-read this plan and confirm the contract decisions still match product expectations.
- [x] Audit host assumptions about callback ordering and progress fields (`IFileSystemCallback` implementations).
- [x] Confirm all `FileInfo` consumers treat it as in-memory only (no persistence).
- [x] Implement phases in order (contract → watcher → host resync → enum → parallel ops).

## Open Questions (resolve before coding the risky parts)

1) **Out-of-order completion**: resolved — host tolerates out-of-order `itemIndex` completion; it derives sequencing/progress from monotonic counters and treats `itemIndex` as informational.
2) **totalBytes growth**: is it acceptable that `totalBytes` may increase during directory copy/move (potentially making derived percent non-monotonic)?
3) **Configuration knobs**: implemented via plugin configuration JSON (see Phase 5b).
4) **Delete parallelism**: separate caps for fast delete vs Recycle Bin delete are retained and are configurable (see Phase 5b).

## References

- Skills:
  - `.github/skills/async-threading/SKILL.md`
  - `.github/skills/error-handling/SKILL.md`
  - `.github/skills/wil-raii/SKILL.md`
  - `.github/skills/plugin-callbacks/SKILL.md`
  - `.github/skills/cpp-modern-style/SKILL.md`
  - `.github/skills/compiler-warnings/SKILL.md`
  - `.github/skills/cpp-build/SKILL.md`
- Related specs:
  - `Specs/FileOperationsSpec.md`
  - `Specs/FolderWindowSpec.md`
  - `Specs/DirectoryInfoCacheSpec.md`
  - `Specs/PluginAPI.md`
