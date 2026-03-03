# FileSystem Code Review & Remediation Plan — Performance, Reliability & Parallelism

**Date:** 2026-02-26  
**Scope:** Built-in `FileSystem` plugin, `FileSystemCurl` plugin, `FileSystem7z` plugin, `FileSystemS3` plugin, plugin interfaces, host-side file operations (`FolderWindow` / popup)  
**Status:** **P0 + P1-1..P1-8 implemented (2026-02-26)** — local + curl `progressStreamId` identity mapping, `FileSystem::Release()` memory ordering, selftest coverage, cancellation-responsive bandwidth sleeps, interface contract documentation updates, debug-only detection for “expected >1 in-flight progress line but observed ≤1”, Curl transfer retry/backoff + easy-handle reuse, Curl recursive copy `continue-on-error`, and explicit shared scheduler shutdown/join at plugin quiet point. Artifacts: `Specs/TestRuns/Tiny/FileOps/2026-02-26_161006/`, `Specs/TestRuns/Tiny/FileOps/2026-02-26_164241/`, `Specs/TestRuns/Tiny/FileOps/2026-02-26_170104/`, `Specs/TestRuns/Tiny/FileOps/2026-02-26_172802/`, `Specs/TestRuns/Tiny/FileOps/2026-02-26_175847/`. **P2-1 implemented (2026-02-27)** — `CopyDirectoryChildrenParallel` streaming queue + cancel checks. Artifacts: `Specs/TestRuns/4cb089111a23/FileOps/2026-02-27_085402/` (baseline), `Specs/TestRuns/4cb089111a23/FileOps/2026-02-27_091634/` (after). **P2-2..P2-5 + P2-8..P2-13 implemented (2026-02-27)** — 7z indexing lock reduction + extract progress/cancel + hardening, reparse retarget perf, host per-item in-flight LRU eviction, and S3 delete callbacks + `DeleteObjects` batching + per-mode capabilities. Artifacts: `Specs/TestRuns/4cb089111a23/FileOps/2026-02-27_105039/` (FileOps selftest pass), `Specs/TestRuns/4cb089111a23/FileOps/2026-02-27_141202/` (FileOps selftest pass; adds Phase 15 7z read/seek smoke on a >32MB entry).

This document is intentionally “decision complete”: it captures **validated discovery**, **what to change**, **where to change it**, and **how to validate** so the next session can directly implement.

## Implementation Checklist (execute top-to-bottom)
This checklist is the “do this next” index for every improvement in this document. We’ll implement it step-by-step in order.

### Baseline (before any changes)
- [x] Tiny baseline captured under `Specs/TestRuns/Tiny/FileOps/<timestamp>/` (baseline was captured on a different machine).
- [x] Baseline captured under `Specs/TestRuns/4cb089111a23/FileOps/2026-02-27_085402/` (baseline for this machine).
- [x] Build Debug.
- [x] Run `RedSalamander.exe --fileops-selftest` and save output (baseline).
- [x] Run the commands in **Next-Session Implementation Commands** to confirm all edit sites are still present (baseline).
- [ ] Local→local repro: set `copyMoveMaxConcurrency = 4`, copy a directory with multiple large files, and observe popup in-flight lines (baseline).
- [ ] FTP/SFTP baseline: copy 50 small files and record total time + any “dead-air” between files.
- [ ] 7z baseline: open/read a large file from a solid archive; test “cancel/close” responsiveness.
- [ ] S3 baseline: attempt delete UI + call `DeleteItem` on a test object; note whether capabilities hide delete.

### Phase 0 (P0) — correctness (must-fix)
- [x] **P0-3:** `Plugins/FileSystem/FileSystem.cpp` — change `FileSystem::Release()` final `fetch_sub(..., std::memory_order_relaxed)` → `std::memory_order_acq_rel`.
- [x] **P0-1:** `Plugins/FileSystem/FileSystem.FileOps.cpp` — replace `progressStreamId = schedulerStreamId % concurrency` with `progressStreamId = schedulerStreamId` at all validated call sites (3).
- [x] **P0-2:** `Plugins/FileSystemCurl/FileSystemCurl.CopyMove.cpp` — replace `progressStreamId = schedulerStreamId % concurrency` with identity at all validated call sites (2).
- [x] **P0-4 (2026-02-28):** `Plugins/FileSystem/FileSystem.FileOps.cpp` — fix potential vector invalidation / use-after-free in `FlattenDeleteDirectoryTree` (avoid holding references into `stack` while pushing).
- [x] **P0-5 (2026-02-28):** `Plugins/FileSystem/FileSystem.FileOps.cpp`, `RedSalamander/FolderWindow.FileOperations.State.cpp` — avoid `wil::CoInitializeEx()` (throwing) inside `noexcept` worker threads; use `wil::CoInitializeEx_failfast()` instead.
- [x] Add/extend `--fileops-selftest` to cover the real production symptom path: **PerItem** copy of a directory where the plugin performs internal parallel child work, and assert the popup shows **> 1** in-flight line (keep a BulkItems variant too).

### Phase 1 (P1) — responsiveness & contract hardening
- [x] **P1-1:** `Plugins/FileSystem/FileSystem.FileOps.cpp` — slice bandwidth-limit sleeps in `CopyProgressRoutine` and check cancel between slices.
- [x] **P1-2:** `Common/PlugInterfaces/FileSystem.h` — document `GetDirectorySize()` callback lifetime (“no callbacks after return”).
- [x] **P1-3:** `Common/PlugInterfaces/FileSystem.h` — document `IFilesInformation::GetBuffer()` stability for the object lifetime (no async mutation after return).
- [x] **P1-4:** `RedSalamander/FolderWindow.FileOperations.State.cpp` — debug-only detection/logging for “configured concurrency > 1 but ≤ 1 in-flight line for an extended interval”.
- [x] **P1-5:** `Plugins/FileSystemCurl/FileSystemCurl.Shared.cpp` — verify connection reuse behavior (CURLSH connect-sharing already exists) and improve if needed (persistent per-worker easy handles / tuning).
- [x] **P1-6:** `Plugins/FileSystemCurl/*` — implement retry/backoff for transient transport failures (with cancel checks) where safe.
- [x] **P1-7:** `Plugins/FileSystemCurl/FileSystemCurl.CopyMove.cpp` — propagate `FILESYSTEM_FLAG_CONTINUE_ON_ERROR` through `CopyDirectoryRecursive`.
- [x] **P1-8:** `Plugins/FileSystem/*` + `Plugins/FileSystemCurl/*` — decide and implement explicit scheduler shutdown/ownership strategy for plugin unload safety (and document host quiet-point expectations).

### Phase 2 (P2) — performance scaling & robustness
- [x] **P2-1:** `Plugins/FileSystem/FileSystem.FileOps.cpp` — refactor `CopyDirectoryChildrenParallel` to bounded producer/consumer (enumerator + workers) with cancellation on both sides. (Validated: `Specs/TestRuns/4cb089111a23/FileOps/2026-02-27_091634/`.)
- [x] **P2-2:** `Plugins/FileSystem7z/FileSystem7z.cpp` — move `EnsureIndex()` heavy work outside `_stateMutex` (build temp, publish/swap if still current).
- [x] **P2-3:** `Plugins/FileSystem7z/FileSystem7z.cpp` — implement 7z open/extract progress callbacks (throttled) and decide how they map to host UI.
- [x] **P2-4:** `Plugins/FileSystem7z/FileSystem7z.cpp` — cap/validate `numItems` before `reserve()`/iteration.
- [x] **P2-5:** `Plugins/FileSystem/FileSystem.FileOps.cpp` — reduce reparse retarget allocations (measure before/after).
- [ ] **P2-6:** `Plugins/FileSystem7z/FileSystem7z.cpp` — evaluate `_extractCv.notify_all()` and pipe copy strategy; only change after measurement.
- [ ] **P2-7:** `Plugins/FileSystem7z/FileSystem7z.cpp` — evaluate backward seek restart behavior; consider caching/segmentation; measure.
- [x] **P2-8 (optional):** `RedSalamander/FolderWindow.FileOperations.State.cpp` — improve `_perItemInFlightCalls` overflow replacement/logging.
- [x] **P2-9:** `Plugins/FileSystem7z/FileSystem7z.cpp` — improve extract stop/cancel responsiveness (return `E_ABORT` from extract callbacks when stop requested).
- [x] **P2-10:** `Plugins/FileSystem7z/FileSystem7z.cpp` — move `GetDirectorySize` callbacks outside `_stateMutex` (avoid deadlock risk).
- [x] **P2-11:** `Plugins/FileSystem7z/FileSystem7z.cpp` — remove redundant `sort+unique` in `GetEntriesForDirectory`.
- [x] **P2-12:** `Plugins/FileSystemS3/*` — implement delete callbacks + batch `DeleteItems` via `DeleteObjects`.
- [x] **P2-13:** `Plugins/FileSystemS3/FileSystemS3.h` — fix capabilities JSON (`"delete": true`).

### Phase 3 (P3) — long-term debt (stage after P0–P2)
- [x] **P3 architecture:** Callback lifetime model + ABI versioning (`sizeBytes`) + debug guards (keep behavior identical). Implemented: `Specs/Plans/Done/RFC_FileSystem_P3Architecture.md` (Validated: `Specs/TestRuns/4cb089111a23/FileOps/2026-03-01_145535/`, `Specs/TestRuns/4cb089111a23/SelfTest/2026-03-01_141839/`.)
- [x] **P3-1:** `Plugins/FileSystem7z/FileSystem7z.cpp` — replace `swscanf_s` parsing + temporary allocation in time parsing (style/perf).
- [x] **P3-2:** `Plugins/FileSystem7z/FileSystem7z.cpp` — harden spool growth / extracted-byte bounds for corrupted archives.
- [x] **P3-3:** `Plugins/FileSystem7z/FileSystem7z.cpp` — convert extract thread to `std::jthread` + `stop_token`.
- [x] **P3-4:** `Plugins/FileSystemCurl/FileSystemCurl.CopyMove.cpp` — parallelize `MoveItems` (with destructive-op safety).

### Final verification gates
- [x] Re-run `RedSalamander.exe --fileops-selftest` (including new assertions).
- [ ] Manual acceptance (local→local): popup shows multiple stable in-flight lines; no “teleporting”; cancel responsive under speed limit.
- [ ] Manual acceptance (FTP/SFTP): reduced dead-air between files; transient drop triggers at least one retry attempt.
- [ ] Manual acceptance (FTP continue-on-error): subtree copy continues after a single file failure when `FILESYSTEM_FLAG_CONTINUE_ON_ERROR` is set.
- [ ] Manual acceptance (7z): cancel/close completes quickly for large solid-archive reads.
- [ ] Manual acceptance (S3): delete is available when it should be; delete operations are cancellable/observable; bulk delete uses batching.

---

## TL;DR (Validated)

### P0: Fix “copy progress lagging / missing callbacks” for local→local copy/move
The primary cause of the reported symptom is **`progressStreamId` collisions** during parallel copy/move:

- Host tracks popup in-flight lines by **`(cookie, progressStreamId)`**.
- UI commands currently run file ops in **`ExecutionMode::PerItem`** with a non-null stack-allocated `PerItemCallbackCookie`. The `BulkItems` path (which passes `cookie == nullptr`) is exercised by `--fileops-selftest` and must stay correct.
- The **real collision scenario** occurs within a **single recursive directory copy**: the plugin spawns multiple parallel workers that all share the **same cookie pointer** (from the caller). Those workers compute `progressStreamId` as **`schedulerStreamId % concurrency`**, which can collide when two workers are congruent mod `concurrency` (e.g. `1 % 4 == 5 % 4`).
- Collisions overwrite UI lines within the same directory copy → appears as "callbacks not sent", "progress jumps", or "lag".

**Fix:** collision-free mapping (identity): `progressStreamId = schedulerStreamId`, plus regression guards.

**Note on P0-2 (FileSystemCurl):** Still matters whenever configured `concurrency` is less than the worker pool size (e.g. `concurrency == 2`); any worker can pick up in-flight items, so modulo mapping can collide even with an 8-worker scheduler.

### Also P0: Fix real ref-count synchronization bug
`Plugins/FileSystem/FileSystem.cpp` has `FileSystem::Release()` using `memory_order_relaxed` before `delete this` → switch to `memory_order_acq_rel`.

---

## Validated Discovery (2026-02-26)

### Host: how progress is tracked
**File:** `RedSalamander/FolderWindow.FileOperations.State.cpp`

- UI commands use `ExecutionMode::PerItem` with a **non-null** stack-allocated `PerItemCallbackCookie`. `--fileops-selftest` exercises `ExecutionMode::BulkItems` (cookie is `nullptr`), so both modes must remain correct.
- The host maintains **two separate** fixed-size arrays (size 8) for progress display:
  - `_perItemInFlightCalls`: keyed by **cookie pointer identity** — tracks aggregate byte/item counters. On overflow, **overwrites `.back()` silently** (no LRU).
  - `_inFlightFiles` (`InFlightFileProgress`): keyed by **`(cookieKey, progressStreamId)` pair** — drives per-file popup lines. On overflow, **replaces oldest entry** (proper LRU by `lastUpdateTick`).
- These are different arrays with different overflow strategies; P2-8 targets `_perItemInFlightCalls` specifically.

### Built-in FileSystem plugin: scheduler vs. job concurrency
**File:** `Plugins/FileSystem/FileSystem.FileOps.cpp`

- `SharedFileOpsJobScheduler` creates up to **8** worker threads (`kMaxWorkers = 8`).
- A job’s `maxConcurrency` (default 4) limits *in-flight items*, not the scheduler’s worker IDs.
- Several worker contexts compute `progressStreamId = schedulerStreamId % concurrency` → collisions occur whenever two simultaneously in-flight workers are congruent mod `concurrency` (e.g. worker 1 and worker 5 when `concurrency == 4`).

### FileSystemCurl plugin: same pattern
**File:** `Plugins/FileSystemCurl/FileSystemCurl.CopyMove.cpp`

- Curl scheduler uses up to **8** worker threads.
- Worker code also uses modulo mapping.

### Interface contract: “no concurrent callbacks” does not forbid parallel work
**File:** `Common/PlugInterfaces/FileSystem.h`

- Plugins may do parallel work; they must **serialize callback invocations**.
- `progressStreamId` is still required to differentiate progress streams even when callbacks are serialized.

---

## Root Cause Analysis — Copy “lagging” / callbacks “missing” (Validated)

### Primary cause: `progressStreamId` collisions overwrite host UI lines
Within a single recursive directory copy, the plugin spawns multiple parallel workers that share the **same cookie pointer** (from the host's `PerItemCallbackCookie`). Collisions collapse multiple workers into a single in-flight line. Typical symptoms:

- fewer in-flight lines than expected (often "stuck at 1")
- line "teleporting" between files
- progress appearing to stall or jump

Concrete example with default `concurrency == 4` and an 8-worker scheduler:
- worker 1 emits `(cookie, 1 % 4)` = `(cookie, 1)`
- worker 5 emits `(cookie, 5 % 4)` = `(cookie, 1)`
→ both streams overwrite the same host in-flight line because `progressStreamId` is the only discriminator within the same cookie.

### Separate contributors (not the missing-line root cause)
- Bandwidth limiting (`Sleep`) reduces throughput when a speed limit is enabled (design improvement).
- Directory-parallel copy currently buffers all child names before starting work (memory + latency for huge directories).

---

## Remediation Plan (Complete)

### Phase 0 — Must-fix (P0)

#### P0-1: Fix `progressStreamId` mapping in built-in FileSystem (local→local)
**File:** `Plugins/FileSystem/FileSystem.FileOps.cpp`

Replace every mapping of the form:

```cpp
context.progressStreamId = (concurrency > 0) ? (schedulerStreamId % static_cast<uint64_t>(concurrency)) : 0;
```

with:

```cpp
context.progressStreamId = schedulerStreamId;
```

**Known locations (validated):**
- `CopyDirectoryChildrenParallel(...)` worker context setup
- `FileSystem::CopyItems(...)` parallel job worker context setup
- `FileSystem::MoveItems(...)` parallel job worker context setup

**Regression guard (compile-time):**
- Put the mapping behind a helper (e.g. `MapProgressStreamId(schedulerStreamId)`) and add a `static_assert` that it is injective for the scheduler worker ID range (0..7).
- Prefer a behavioral guard too: add/extend `--fileops-selftest` to run a bulk copy with `copyMoveMaxConcurrency == 4` and verify the host observes >1 in-flight progress line (this catches regressions regardless of implementation detail).

#### P0-2: Fix `progressStreamId` mapping in FileSystemCurl
**File:** `Plugins/FileSystemCurl/FileSystemCurl.CopyMove.cpp`

Replace:

```cpp
const uint64_t progressStreamId =
    (concurrency > 0) ? (schedulerStreamId % static_cast<uint64_t>(concurrency)) : 0;
```

with:

```cpp
const uint64_t progressStreamId = schedulerStreamId;
```

**Regression guard (compile-time):**
- Add an injectivity `static_assert` for Curl scheduler worker IDs (0..3).

#### P0-3: Fix reference-count memory ordering in `FileSystem::Release`
**File:** `Plugins/FileSystem/FileSystem.cpp`

- Change final `fetch_sub()` from `std::memory_order_relaxed` to `std::memory_order_acq_rel` before `delete this`.
- Note: this may appear benign on x64, but is a real correctness issue on weaker memory models (ARM64 is a target).

---

### Phase 1 — Responsiveness & contract hardening (P1)

#### P1-1: Make bandwidth limiting cancellation-responsive (local FileSystem)
**File:** `Plugins/FileSystem/FileSystem.FileOps.cpp` (`CopyProgressRoutine`)

- Replace long `Sleep(sleepMs)` with sliced sleeps (e.g. 25–50ms slices).
- Between slices: check cancel/stop flags and return `PROGRESS_CANCEL` quickly.

#### P1-2: Clarify `GetDirectorySize` callback lifetime contract
**File:** `Common/PlugInterfaces/FileSystem.h` (comment-only)

- Explicitly state that `GetDirectorySize()` callbacks are only invoked during the call and never after it returns (host may pass stack cookies).

#### P1-3: Clarify `IFilesInformation` buffer lifetime contract
**File:** `Common/PlugInterfaces/FileSystem.h` (comment-only)

- Document that the buffer returned from `IFilesInformation::GetBuffer()` is stable for the lifetime of the `IFilesInformation` object and will not be mutated asynchronously after `ReadDirectoryInfo()` returns.

#### P1-4: Debug diagnostics for progress-stream collisions (host-side, debug-only)
**File:** `RedSalamander/FolderWindow.FileOperations.State.cpp`

- Add a debug-only warning when an operation is configured for concurrency > 1 but observed in-flight file line count stays ≤ 1 for an extended interval (covers both `cookie == nullptr` BulkItems and same-cookie internal-parallel collisions).

#### P1-5: FileSystemCurl — connection reuse across transfers
**File:** `Plugins/FileSystemCurl/FileSystemCurl.Shared.cpp`

Implemented: `CurlDownloadToFile` / `CurlUploadFromFile` reuse a thread-local `CURL*` easy handle (via `curl_easy_reset()` between transfers) to reduce per-transfer setup overhead.

Important nuance (code-backed): `ApplyCommonCurlOptions()` already attaches a process-wide `CURLSH*` share handle with `CURL_LOCK_DATA_CONNECT`, `CURL_LOCK_DATA_SSL_SESSION`, and `CURL_LOCK_DATA_DNS`, so connection reuse is already attempted across easy handles.

**Improvement path (if “dead-air” persists):**
- Measure whether connections are actually being reused (enable curl verbose in a debug build or add targeted perf logs around `curl_easy_perform`).
- If reuse is not effective (or per-transfer setup overhead dominates), keep a small pool of persistent `CURL*` easy handles (e.g. per scheduler worker) and reuse them via `curl_easy_reset()` between transfers.

**Validation:** batch-copy 50 small files over FTP/SFTP; compare total time and per-file “gaps” before/after.

#### P1-6: FileSystemCurl — retry logic for transient transport failures
**Files:** `Plugins/FileSystemCurl/FileSystemCurl.Shared.cpp`, `Plugins/FileSystemCurl/FileSystemCurl.CopyMove.cpp`

Implemented: `CurlDownloadToFile` / `CurlUploadFromFile` now retry transient transport failures once (2 attempts total) with a short backoff. The existing `ResolveLocationWithAuthRetry(...)` helper still only targets **Connection Manager authentication** scenarios; this change covers transport-level errors (timeout, connection reset, recv/send failure). Cancel is checked between retries when a progress context exists.

**Fix:** Add at least one retry with exponential backoff for transient `CURLE_` error codes: `CURLE_OPERATION_TIMEDOUT`, `CURLE_RECV_ERROR`, `CURLE_SEND_ERROR`, `CURLE_GOT_NOTHING`, `CURLE_COULDNT_CONNECT` (on retry only). Check cancel flag between retries.

#### P1-7: FileSystemCurl — `continue-on-error` not propagated into recursive directory copy
**File:** `Plugins/FileSystemCurl/FileSystemCurl.CopyMove.cpp` (`CopyDirectoryRecursive`)

Implemented: `CopyDirectoryRecursive` respects `FILESYSTEM_FLAG_CONTINUE_ON_ERROR`: per-child failures are recorded and the copy continues, returning `HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY)` if any child failed. Cancellation and authentication failures still abort immediately.

**Fix:** Thread the `continueOnError` flag through `CopyDirectoryRecursive` and its recursive calls. On per-file failure with the flag set, record the error, continue with the next file, and return `HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY)` if any child failed.

#### P1-8: Static scheduler lifetime vs. plugin unload safety
**Files:** `Plugins/FileSystem/FileSystem.FileOps.cpp` (`SharedFileOpsJobScheduler`), `Plugins/FileSystem/FileSystem.Watch.cpp` (`FileSystem::~FileSystem`), `Plugins/FileSystemCurl/FileSystemCurl.CopyMove.cpp` (`SharedCopyMoveJobScheduler`), `Plugins/FileSystemCurl/FileSystemCurl.Shared.cpp` (`FileSystemCurl::Release`)

Implemented: both plugins now expose an explicit “shutdown and join” path for their shared copy/move schedulers and invoke it when the **last** plugin instance is being destroyed (host quiet point). This requests stop, wakes workers, and joins `std::jthread`s before DLL unload, avoiding reliance on DLL static destruction order. The schedulers are restartable if a new instance is created later (workers are recreated on demand).

**Fix:** Keep schedulers in named internal namespaces and provide `ShutdownShared*JobScheduler()` helpers called when the last instance is destroyed (`FileSystem::~FileSystem`, `FileSystemCurl::Release`).

---

### Phase 2 — Performance scaling & robustness (P2)

#### P2-1: Stream directory-parallel copy instead of buffering all children
**File:** `Plugins/FileSystem/FileSystem.FileOps.cpp` (`CopyDirectoryChildrenParallel`)

Implemented (2026-02-27): `CopyDirectoryChildrenParallel` now uses a bounded producer/consumer queue (one enumerator, N workers) so large directories don’t require buffering all child names before starting copy work.

- Preserves semantics (reparse policy, conflicts, serialized callbacks).
- Cancellation stops both sides (workers **and** enumerator) so enumeration does not continue buffering after the operation is cancelled.

Validation artifacts: `Specs/TestRuns/4cb089111a23/FileOps/2026-02-27_091634/` (P2-1), `Specs/TestRuns/4cb089111a23/FileOps/2026-02-27_105039/` (post-P2 FileOps selftest pass).

#### P2-2: 7z indexing should not hold `_stateMutex` for full index build
**File:** `Plugins/FileSystem7z/FileSystem7z.cpp` (`EnsureIndex`)

Implemented (2026-02-27): `EnsureIndex()` now snapshots `(_archivePath, _password)` under `_stateMutex`, builds the index outside the lock, then publishes (swap) under lock only if inputs are still current.

Build validation: `.\build.ps1 -Configuration Debug -ProjectName FileSystem7z`.

#### P2-3: Implement 7z extraction progress callbacks (currently stubbed)
**File:** `Plugins/FileSystem7z/FileSystem7z.cpp`

Implemented (2026-02-27): `SevenZipItemFileReader::ExtractCallback` now implements `IProgress::SetTotal()` / `SetCompleted()` and returns `E_ABORT` when stop is requested (atomic flag) for faster cancel/close responsiveness.

Build validation: `.\build.ps1 -Configuration Debug -ProjectName FileSystem7z`.

#### P2-4: Cap/validate 7z `numItems` before reserving/iterating
**File:** `Plugins/FileSystem7z/FileSystem7z.cpp` (`BuildIndex`)

Implemented (2026-02-27): `BuildIndex(...)` caps `numItems` (1,000,000) and returns `HRESULT_FROM_WIN32(ERROR_INVALID_DATA)` when exceeded to avoid reserve/iteration on malicious archives.

Build validation: `.\build.ps1 -Configuration Debug -ProjectName FileSystem7z`.

#### P2-5: Reduce allocations in reparse-point retargeting
**File:** `Plugins/FileSystem/FileSystem.FileOps.cpp` (`TryRetargetPathIntoDestination`)

Implemented (2026-02-27): `TryRetargetPathIntoDestination` now trims via `std::wstring_view` and only allocates when slash normalization is needed, reducing per-call temporary `std::wstring` churn.

Build validation: `.\build.ps1 -Configuration Debug -ProjectName FileSystem`.

#### P2-6: Review 7z extract streaming wakeups and pipe design
**File:** `Plugins/FileSystem7z/FileSystem7z.cpp` (extract pipe / `_extractCv`)

- Consider `notify_one()` when only one waiter needs waking.
- Re-evaluate circular-buffer wrap copy cost vs. complexity; measure before changing.

#### P2-7: Reduce backward-seek cost in 7z file reads
**File:** `Plugins/FileSystem7z/FileSystem7z.cpp` (`RestartExtractStreaming`)

- Current behavior restarts extraction on backward seeks; consider caching/segmenting to support common random-access patterns.

#### P2-8 (optional): Improve host per-item in-flight call tracking overflow behavior
**File:** `RedSalamander/FolderWindow.FileOperations.State.cpp`

- **Clarification:** the host uses **two separate** progress arrays with **different overflow strategies**:
  - `_perItemInFlightCalls` (aggregate byte/item counters, keyed by `cookie`): on overflow, **overwrites `.back()` silently** — no LRU, no logging. Lost progress for the evicted cookie until reconciled at item completion.
- `_inFlightFiles` (popup per-file lines, keyed by `(cookie, progressStreamId)`): on overflow, **replaces oldest entry** (proper LRU by `lastUpdateTick`). This one is already reasonable.
- **Fix target:** `_perItemInFlightCalls` — prefer an "oldest/least-recently-updated" replacement strategy (matching `_inFlightFiles` behavior) and/or log a debug warning on eviction.

Implemented (2026-02-27): `_perItemInFlightCalls` now tracks `lastUpdateTick` and uses “oldest tick” eviction (LRU) when full, with a debug warning when an eviction happens.

Build validation: `.\build.ps1 -Configuration Debug -ProjectName RedSalamander`.

#### P2-9: FileSystem7z — extraction stop/cancel responsiveness
**File:** `Plugins/FileSystem7z/FileSystem7z.cpp` (`ExtractThreadMain`, extract pipe)

Implemented (2026-02-27): `ExtractCallback::SetTotal()` / `SetCompleted()` are no longer stubbed; they now track progress and return `E_ABORT` when `_extractStopRequested` is set (atomic), providing a cheap cancel checkpoint during long decompression phases.

Build validation: `.\build.ps1 -Configuration Debug -ProjectName FileSystem7z`.

#### P2-10: FileSystem7z — `GetDirectorySize` calls callback under `_stateMutex`
**File:** `Plugins/FileSystem7z/FileSystem7z.cpp` (`GetDirectorySize`)

Implemented (2026-02-27): `GetDirectorySize` snapshots the needed entry data under `_stateMutex`, releases the lock, then invokes callbacks outside the mutex (and respects callback HRESULTs for progress/cancel).

Build validation: `.\build.ps1 -Configuration Debug -ProjectName FileSystem7z`.

#### P2-11: FileSystem7z — redundant `sort+unique+erase` on children lists
**File:** `Plugins/FileSystem7z/FileSystem7z.cpp` (`BuildIndex`, `GetEntriesForDirectory`)

Implemented (2026-02-27): removed redundant `sort+unique+erase` from `GetEntriesForDirectory` (children are already sorted+unique from index build).

Build validation: `.\build.ps1 -Configuration Debug -ProjectName FileSystem7z`.

#### P2-12: FileSystemS3 — `DeleteItem` ignores all callbacks; no batch delete
**File:** `Plugins/FileSystemS3/FileSystemS3.Directory.cpp`

Implemented (2026-02-27): `DeleteItem` now checks cancel and reports progress + item completion via `IFileSystemCallback`. `DeleteItems` now uses S3 `DeleteObjects` batching (up to 1000 keys per request), with per-item completion reporting and cancel/progress checkpoints.

Build validation: `.\build.ps1 -Configuration Debug -ProjectName FileSystemS3`.

#### P2-13: FileSystemS3 — capabilities declare `delete: false` but `DeleteItem` works
**File:** `Plugins/FileSystemS3/FileSystemS3.h` (capabilities JSON)

Implemented (2026-02-27): capabilities are now per-mode — S3 declares `"delete": true`, while S3 Table remains `"delete": false`.

---

### Phase 3 — Interface/architecture changes & hardening (P3, long-term)
Stage only after P0/P1 are stable.

- Callback lifetime model: COM-ify callbacks (IUnknown) or introduce explicit registration tokens with "no callback after unregister" guarantees.
- Interface versioning and ABI stability fields for structs.
- Reduce API duplication between single-item and batch operations while preserving behavior.

#### P3-1: FileSystem7z — avoid `swscanf_s` + temporary allocation
**File:** `Plugins/FileSystem7z/FileSystem7z.cpp` (`TryParseModifiedLocalTime`, line ~1617)

Uses `swscanf_s(std::wstring(text).c_str(), ...)` which forces a temporary allocation and adds locale/format fragility. Prefer manual parsing (fixed-format) or `std::from_chars` for date components.

Implemented (2026-02-27): replaced `swscanf_s` parsing and temporary allocation with fixed-format parsing over `std::wstring_view` (no allocation) and strict range validation.

#### P3-2: FileSystem7z — harden `_spool` growth / extracted-byte bounds
**File:** `Plugins/FileSystem7z/FileSystem7z.cpp` (`WriteExtractBytes`, line ~3431)

In in-memory spool mode, `WriteExtractBytes` appends extracted bytes into `_spool` and currently does not validate extracted byte totals against the expected `_fileSizeBytes`. Corrupt archives or unexpected extractor behavior can cause pathological growth or hangs.

**Fix:** Abort extraction if extracted bytes exceed the expected `_fileSizeBytes` (or if in-memory spool growth exceeds a hard cap) to prevent runaway allocations or hangs on corrupted archives; return an error (`ERROR_INVALID_DATA`/`E_ABORT`) to stop extraction.

Implemented (2026-02-27): added extracted-byte bounds validation in both pipe + in-memory spool paths to abort extraction on overrun (prevents runaway allocations / hangs on corrupted archives).

#### P3-3: FileSystem7z — raw `std::thread` instead of `std::jthread`
**File:** `Plugins/FileSystem7z/FileSystem7z.cpp` (`SevenZipItemFileReader`, line ~3479)

The extract thread is created as a raw `std::thread`, requiring manual stop-flag management and explicit `join()`. Per project guidelines, prefer `std::jthread` with `stop_token` for automatic cancellation and join-on-destruction.

Implemented (2026-02-27): convert `_extractThread` to `std::jthread` and propagate cancellation via `stop_token` (with existing stop flag still gating extract callbacks + waits).

#### P3-4: FileSystemCurl — `MoveItems` is entirely sequential
**File:** `Plugins/FileSystemCurl/FileSystemCurl.CopyMove.cpp` (`MoveItems`, line ~1445)

`CopyItems` dispatches work through the parallel `SharedCopyMoveJobScheduler`, but `MoveItems` processes items in a simple sequential loop. This asymmetry means batch moves are significantly slower than batch copies for no documented reason. Consider parallelizing `MoveItems` using the same scheduler, with appropriate safety guards for destructive operations.

Implemented (2026-02-27): `MoveItems` now uses `SharedCopyMoveJobScheduler` to process items in parallel (matching `CopyItems` behavior) while preserving per-item completion reporting + continue-on-error semantics.

---

## Validation Plan

### Automated (preferred)
- Debug build: `RedSalamander.exe --fileops-selftest`

### Manual acceptance (local→local)
1) Set FileSystem plugin `copyMoveMaxConcurrency` to 4.
2) Copy a directory with multiple large files.
3) Pass criteria:
   - popup shows multiple stable in-flight lines
   - no “teleporting” single-line effect
   - cancel remains responsive even with a speed limit enabled

### Manual acceptance (FTP/SFTP — P1-5/P1-6)
1) Batch-copy 50 small files (< 1KB each) over FTP.
2) Pass criteria:
   - total time ≥2× faster after connection reuse (no dead-air between files)
   - on simulated network drop mid-transfer: at least one retry attempt before failure

### Manual acceptance (FTP continue-on-error — P1-7)
1) Copy a directory tree with a deliberately inaccessible file in the middle.
2) Set `FILESYSTEM_FLAG_CONTINUE_ON_ERROR`.
3) Pass criteria: remaining files in the subtree are copied; inaccessible file reported as error but does not abort siblings.

### Manual acceptance (7z cancellation — P2-9)
1) Open a large file (> 100MB) from a solid 7z archive.
2) Immediately request cancellation.
3) Pass criteria: cancellation completes within ~2 seconds, not after full extraction.

### Manual acceptance (S3 delete — P2-12)
1) Select 100+ objects in an S3 bucket.
2) Delete all.
3) Pass criteria: operation uses batch `DeleteObjects` API (verify via S3 API call count if possible).

---

## Next-Session Implementation Commands
Use these to quickly locate the P0/P1/P2 edit sites:

```bash
# P0-1/P0-2: progressStreamId modulo mapping
rg -n "schedulerStreamId %\s*static_cast<uint64_t>\(concurrency\)" Plugins/FileSystem/FileSystem.FileOps.cpp
rg -n "schedulerStreamId %\s*static_cast<uint64_t>\(concurrency\)" Plugins/FileSystemCurl/FileSystemCurl.CopyMove.cpp

# P0-3: relaxed memory order in Release
rg -n "fetch_sub\(1, std::memory_order_relaxed\)" Plugins/FileSystem/FileSystem.cpp

# P1-1: Sleep calls in bandwidth limiting
rg -n "Sleep\(" Plugins/FileSystem/FileSystem.FileOps.cpp

# P1-5: Curl handle creation (connection reuse)
rg -n "curl_easy_init" Plugins/FileSystemCurl/FileSystemCurl.Shared.cpp

# P1-6: Curl error handling (retry candidates)
rg -n "CURLE_" Plugins/FileSystemCurl/FileSystemCurl.Shared.cpp

# P1-7: Recursive copy without continue-on-error
rg -n "CopyDirectoryRecursive" Plugins/FileSystemCurl/FileSystemCurl.CopyMove.cpp

# P1-8: Static scheduler singletons
rg -n "static SharedFileOpsJobScheduler|static SharedCopyMoveJobScheduler" Plugins/

# P2-9: 7z extract thread cancellation
rg -n "SetCompleted|SetTotal|_extractStopRequested" Plugins/FileSystem7z/FileSystem7z.cpp

# P2-10: GetDirectorySize under mutex
rg -n "GetDirectorySize" Plugins/FileSystem7z/FileSystem7z.cpp

# P2-11: Redundant sort+unique+erase
rg -n "std::sort|std::unique" Plugins/FileSystem7z/FileSystem7z.cpp

# P2-12: S3 DeleteItem callback usage
rg -n "DeleteItem|DeleteItems|maybe_unused.*callback" Plugins/FileSystemS3/

# P2-13: S3 capabilities
rg -n "delete.*false|delete.*true" Plugins/FileSystemS3/FileSystemS3.h

# P3-1: swscanf_s in 7z
rg -n "swscanf_s" Plugins/FileSystem7z/FileSystem7z.cpp

# P3-3: raw std::thread in 7z
rg -n "std::thread" Plugins/FileSystem7z/FileSystem7z.cpp
```

---

## Previous Draft Issue IDs — Status After Validation (2026-02-26)
This section exists to challenge the earlier draft and prevent bad fixes.

| Draft ID | Previous claim (short) | Status | Replacement / Note |
|---|---|---:|---|
| CRIT-01 | multiple Release() relaxed order | Partially valid | Only `FileSystem::Release()` is relaxed (P0-3). |
| CRIT-02 | null callback deref in watch | Invalidated | Watch checks `_callback` and Stop drains callbacks. |
| CRIT-03 | callback lifetime w/o refcount | Design risk | Real risk for async callbacks (watch); not proven crash. |
| CRIT-04 | GetDirectorySize cookie stack lifetime | Contract clarification | Document sync-only callbacks (P1-2). |
| HIGH-01 | Sleep blocks copy threads | Valid (design) | Improve cancellation responsiveness (P1-1). |
| HIGH-02 | sleep while holding callback lock | Invalidated | Sleep is outside callback mutex; callbacks are serialized by contract. |
| HIGH-03 | OperationContext shared race | Invalidated | Parallel paths use per-worker contexts. |
| HIGH-04 | host progress cookie data races | Invalidated | Mutations are under `_progressMutex`; callbacks serialized. |
| HIGH-05 | 7z index holds mutex long | Valid | Keep as P2-2. |
| HIGH-06 | host in-flight array concurrent modification | Invalidated | Mutations are under `_progressMutex`. |
| MED-01 | bandwidth calc overflow | Invalidated | Code guards overflow via `maxSafeBytes`. |
| MED-02 | interface contract contradiction | Invalidated | Parallel work + serialized callbacks is compatible. |
| MED-03 | 7z extract progress missing | Valid | Implement callbacks (P2-3). |
| MED-04 | NtPathToWin32Path duplicated check | Invalidated | Handles both `\\??\\` and `\\\\?\\` prefixes. |
| MED-05 | per-item in-flight overflow corrupts progress | Partially valid | Best-effort UI cap; consider better replacement/logging (P2-8). |
| MED-06 | 7z numItems unbounded | Valid (hardening) | Clamp/validate `numItems` (P2-4). |
| MED-07 | UnwatchDirectory race | Invalidated | Plugin Stop waits for IO/work callbacks to finish. |
| MED-08 | _lastItemIndex/_lastItemHr races | Invalidated | Updated under `_progressMutex`; no reads without lock found. |
| MED-09 | lock-order deadlock risk | Design note | Current code appears safe; document lock ordering. |
| MED-10 | IFilesInformation buffer volatility | Design note | Clarify immutability/lifetime (P1-3). |
| PERF-01 | reparse retarget allocs | Perf opportunity | Optimize with views (P2-5), measure. |
| PERF-02 | pipe wrap double-memcpy | Perf opportunity | Likely fine; measure before change (P2-6). |
| PERF-03 | progress path string churn | Perf opportunity | Micro-opt; consider only if profiling shows (optional). |
| PERF-04 | notify_all thundering herd | Perf opportunity | Consider notify_one when appropriate (P2-6). |
| PERF-05 | backward seek restarts extraction | Perf opportunity | Consider caching/segmenting (P2-7). |
| INTF-01 | single vs batch duplication | Design | Covered in P3. |
| INTF-02 | GetCapabilities pointer unsafe | Invalidated | Host can copy string; lifetime is documented. |
| INTF-03 | options in/out confusion | Design | Clarify docs; current pattern is workable. |
| INTF-04 | missing ABI versioning | Valid (design) | Covered in P3. |
| INTF-05 | JSON without schema/limits | Design | Add size limits/schema in future (P3). |
