# Compare Directories Engine (Developer)

This page is the developer-level companion to the user-facing [CompareDirectories.md](../CompareDirectories.md) and the architecture summary in [DeveloperGuide.md](../DeveloperGuide.md). It documents the internals of the Compare Directories diff engine: the session/engine object model, the scan and content-compare worker pools, the atomic `_version` coherence model, the posted-message protocol that reaches the UI, and the decision-cache LRU (budget, pinning, eviction). Everything here is grounded in `RedSalamander/CompareDirectoriesEngine.h/.cpp` and the window layer in `RedSalamander/CompareDirectoriesWindow*.cpp`.

The normative contract lives in `Specs/Core/Core_CompareDirectories.md`; this page describes how the current code implements it.

## Overview

The feature splits into a UI-agnostic engine and a window/host layer:

| Layer | File(s) | Responsibility |
|-------|---------|----------------|
| Engine session | `CompareDirectoriesEngine.h/.cpp` (`CompareDirectoriesSession`) | Owns roots, settings, decision cache, scan + content-compare worker pools, versioning |
| Filesystem wrapper | `CompareDirectoriesEngine.cpp` (`CompareDirectoriesFileSystem`) | Per-pane `IFileSystem`/`IInformations` wrapper; cache-only enumeration |
| Window | `CompareDirectoriesWindow.cpp` + `.Internal.h` | Top-level window, banner/menu, splitter, embeds a `FolderWindow` |
| Window parts | `.Options.cpp` / `.Progress.cpp` / `.Menu.cpp` | Options panel, progress UI + File Operations task card, themed menu bar |

A session is created with two base `IFileSystem` instances (left/right), two roots, and a `Common::Settings::CompareDirectoriesSettings`. Each pane is then handed a `CompareDirectoriesFileSystem` produced by `CreateCompareDirectoriesFileSystem(pane, session)`. The window embeds a normal `FolderWindow`, so all enumeration flows through these wrappers.

## Session and decision model

The comparison is directory-oriented: for each relative folder under the roots, the engine computes a `CompareDirectoriesFolderDecision` whose `items` is an ordered `std::map<std::wstring, CompareDirectoriesItemDecision, WStringViewNoCaseLess>` keyed by name. Names are matched with Windows ordinal, case-insensitive semantics (`wil::compare_string_ordinal(..., true)`); trailing spaces/dots are stripped by `NormalizeEntryNameForCompare` so handle-based and `FindFirstFile` backends agree.

> The map is deliberately an ordered `std::map`, never an `unordered_map`: ordinal case-insensitive equality is not consistent with a hash, so a hash map would violate the hash/equality contract.

Each `CompareDirectoriesItemDecision` carries existence flags, a `differenceMask` (a bitmask of `CompareDirectoriesDiffBit`), `isDifferent`, `selectLeft`/`selectRight`, and per-side metadata. The folder decision also precomputes `anyDifferent` and `anyPending` aggregates (via `AnyChildDifferent` / `AnyChildPending`) so hot paths avoid O(n) scans, plus a `pendingContentCompareCount` for elided placeholders (see [Differences-only elision](#differences-only-elision-keepidenticalitems)).

### Difference bits

`CompareDirectoriesDiffBit` (in `CompareDirectoriesEngine.h`) is a `uint32_t` flag enum:

| Bit | Meaning |
|-----|---------|
| `OnlyInLeft` / `OnlyInRight` | Item exists on only one side |
| `TypeMismatch` | Same name, but one side is a file and the other a directory |
| `Size` / `DateTime` / `Attributes` | Metadata criteria differ (when the matching setting is enabled) |
| `Content` | File content differs (confirmed by a content compare) |
| `ContentPending` | A content compare is queued or running, not yet resolved |
| `SubdirAttributes` | Directory attributes differ (`compareSubdirectoryAttributes`) |
| `SubdirContent` | A descendant differs (`compareSubdirectories`) |
| `SubdirPending` | A descendant comparison is still in flight |

`HasFlag(mask, bit)` and the `operator|` overloads in the header keep bit math readable. `ContentPending` and `SubdirPending` are explicitly *not* final differences and must not select an item.

## Phase 1: metadata scan

`ComputeDecisionForFolder` performs the synchronous, per-folder metadata pass. It:

1. Resolves the relative folder to absolute paths on both panes (`ResolveAbsolute`, which understands Win32 vs plugin-path roots via `NavigationLocation`).
2. Enumerates each side with `TryReadDirectoryEntries`, which calls `IFileSystem::ReadDirectoryInfo` directly on the base filesystem.
3. Seeds the item map from the left side (preserving left casing as the key), then folds in the right side.
4. Computes existence / type / size / time / attribute differences according to the enabled settings.
5. For files that still need content comparison (`compareContent` on, same size, content compare supported), enqueues a `ContentCompareJob` and sets `ContentPending`.

> Scans deliberately bypass `DirectoryInfoCache::BorrowDirectoryInfo(..., AllowEnumerate)`. That cache is a global cache sized for interactive browsing and can grow to GiB scale when a subtree scan touches many folders. `TryReadDirectoryEntries` walks the raw `FileInfo` buffer returned by `ReadDirectoryInfo` instead (with full bounds checking on `NextEntryOffset`/`FileNameSize`).

Ignore patterns are split once (`SplitPatternsCapped`, capped at 32 patterns of 128 chars) and matched by `MatchesAnyPattern`. Wildcard matching (`WildcardMatchNoCaseBudgeted`) is budgeted (`WildcardBudgetFor`) to harden against pathological globs — exceeding the budget is treated as "no match".

Enumeration failures are recorded as a failed `HRESULT` on the decision; the folder is then treated as different and the failed decision is **not cached** (it is retried on next access). Missing-path errors (`IsMissingPathError`) are not failures — the side enumerates as empty.

## Phase 2: content compare

When `compareContent` is enabled and both sides expose `IFileSystemIO` (`IsContentCompareSupported`), files with equal known sizes are byte-compared asynchronously.

### Worker pools

Both pools are `std::vector<std::jthread>` created lazily under `_mutex`:

| Pool | Created by | Count |
|------|-----------|-------|
| Scan | `EnsureScanWorkersLocked` | `hardware_concurrency()` (fallback 2), clamped to `1..4` |
| Content compare | `EnsureContentCompareWorkersLocked` | `contentCompareWorkerCount` if non-zero else `hardware_concurrency()`/2, clamped to `1..4` |

Each worker calls `CoInitializeEx(COINIT_MULTITHREADED)` and drops itself to `THREAD_PRIORITY_BELOW_NORMAL` so background diffing never starves the UI thread. Workers wait on condition variables (`_scanCv`, `_contentCompareCv`) and exit when their `std::stop_token` is signalled (the destructor calls `request_stop()` then `notify_all()`).

### Two-tier priority queues

Both subsystems keep a **high** and a **low** queue:

- **Scan** (`_scanQueueHigh` / `_scanQueueLow`): `RequestScanForFolder` (triggered by `ReadDirectoryInfo` on a folder the user is looking at) enqueues `ScanPriority::High`; the recursive subtree walk and `StartScan` enqueue `ScanPriority::Low`. `EnqueueScanLocked` dedups by key via `_scanScheduledKeys`; a folder already scheduled low can get a single high-priority "upgrade" duplicate (tracked by `_scanHighQueuedKeys`).
- **Content** (`_contentCompareQueueHigh` / `_contentCompareQueueLow`): a job inherits the priority of the scan that produced it. `ContentCompareWorker` prefers the high queue but, under sustained high load, takes a low job after `kHighBurstBeforeLow` (8) high jobs so background progress is not starved ("visible-first").

This is the backpressure contract: work for currently visible folders is prioritised over low-priority subtree scanning.

### Bounded content queue (backpressure)

The content queue is bounded: `kContentCompareQueueMaxHighJobs` = 128, `kContentCompareQueueMaxLowJobs` = 896. When `ComputeDecisionForFolder` wants to enqueue and the target queue is full, it waits on `_contentCompareQueueNotFullCv` (and re-checks cancellation each wake), so a scan thread blocks rather than letting the queue grow unbounded. Workers `notify_all()` on this CV after dequeuing.

### Byte comparison

`CompareFileContent` opens both sides via `IFileSystemIO::CreateFileReader`, then streams with two 256 KB buffers:

- If both `GetSize` calls succeed and the sizes differ -> `Different` immediately; both zero -> `Equal` immediately.
- Otherwise it reads, tolerates short reads, and `memcmp`s the overlapping region until a mismatch or EOF.
- If sizes are unknown it streams to EOF on both sides and checks for trailing bytes on either.
- It re-checks cancellation (`stop_token`, `_version`, `_backgroundWorkCancelToken`) inside the read loop and reports progress (throttled) via the supplied callback.

A failed reader open or read returns `Different` (conservative — a file that cannot be read is surfaced rather than silently treated as equal).

### Applying results

On completion the worker writes the boolean result into `_contentCompareCache` (bounded at 16384 entries, evicted in batches of 4096) and pushes a `PendingContentCompareUpdate` into `_pendingContentCompareUpdates[folderKey][entryName]`. The actual mutation of cached decisions happens in `ApplyPendingContentCompareUpdatesLocked`, which:

- Drops stale updates whose `version` no longer matches.
- Skips updates whose queued metadata signature no longer matches the cached item (the file changed while the job ran).
- Clears `ContentPending`, sets `Content`/selection if different, and re-derives the item.
- Decrements `pendingContentCompareCount` (and may synthesize an item) for elided differences-only entries.
- Recomputes `anyDifferent`/`anyPending` and calls `PropagateChildAggregateToAncestorsLocked` so directory `SubdirContent`/`SubdirPending` bits transition correctly up the tree.

Application is driven from the UI thread in bounded batches via `FlushPendingContentCompareUpdatesBudgeted(maxFolders)` and `FlushPendingSubdirUpdatesBudgeted(maxKeys)`, both of which return `true` while more work remains.

### Differences-only elision (`keepIdenticalItems`)

When `keepIdenticalItems` is off, `ComputeDecisionForFolder` prunes identical entries from the cached decision to keep memory bounded on very large folders. Per-file `ContentPending` placeholders are also elided, but counted in `pendingContentCompareCount` so `anyPending` stays accurate and folder-level progress UI still works. Toggling `keepIdenticalItems` changes the cached decision shape, so it forces an invalidation and rescan (see [Coherence](#cache-coherence-the-_version-model)).

## Cache coherence: the `_version` model

The session is fully thread-safe under a single `mutable std::mutex _mutex` plus a set of atomics. The primary coherence mechanism is the atomic `uint64_t _version` (starts at 1):

- It increments on `SetRoots`, `SetSettings` (only when a comparison-affecting field changed), and `Invalidate`.
- Every cached `CompareDirectoriesFolderDecision` is stamped with `version` at computation time. A cache hit is valid only if `decision->version == _version.load()`; otherwise it is recomputed.
- Every `FolderScanJob` and `ContentCompareJob` carries the `version` at enqueue time. Workers re-load `_version` after dequeuing and bail out (without writing results) if it changed. The byte-compare loop also re-checks it.
- On a version bump, `ResetCompareStateLocked` swaps out all queues, in-flight maps, and caches into a `ResetCleanup` struct that is freed off-thread via `ScheduleResetCleanup` (so a Rescan over a huge tree does not stall the UI thread destroying state).

### The cancel token

A second atomic, `_backgroundWorkCancelToken`, supports a responsive Cancel that does not invalidate cached decisions. `SetBackgroundWorkEnabled(false)` increments it, clears all queued/in-flight work, and zeroes the progress counters; workers compare it the same way they compare `_version`. The window's banner **Cancel** uses this for snappy stop without a full version bump. Background work stays paused until the next explicit Rescan / Options -> OK re-enables it.

### In-flight stamps

`_scanInFlightKeys` and `_contentCompareInFlight` store an `(version, cancelToken)` stamp alongside the key, not just the key. When a worker finishes it only clears the in-flight entry if the stamp still matches what it inserted. This prevents a stale completion from a restarted run erasing the *new* run's in-flight entry for the same key, which would otherwise undercount `_scanActiveScans` and confuse idle/drain detection.

### `_uiVersion`

A separate, mutex-guarded `uint64_t _uiVersion` tracks the last state the UI has observed. It is bumped on invalidation and whenever decisions change, so the window can skip redundant pane re-enumerations when `_uiVersion` has not moved. It is independent of `_version` (which is purely a cache-coherence counter).

### Invalidation after file operations

After copy/move/delete in compare mode, the window calls `InvalidateForAbsolutePath(absolutePath, includeSubtree)`. It converts the path to a relative path under each root (`TryMakeRelative`), then `InvalidateForRelativePathLocked` erases the matching cache entry plus, for a subtree, all descendant entries (using the ordered map's `lower_bound` prefix range), evicts matching content-compare cache entries, walks ancestors clearing their cached decisions, and bumps `_uiVersion`.

## Decision-cache LRU

The decision cache is `_cache` (folder key -> `shared_ptr<const CompareDirectoriesFolderDecision>`), backed by an LRU and a byte budget so a deep recursive scan cannot grow memory without bound.

### Structures

| Member | Role |
|--------|------|
| `_cache` | The decisions, keyed by `MakeCacheKey(relativeFolder)` (`.` for the root, generic `/`-separated path otherwise) |
| `_decisionCacheLru` | `std::list<std::wstring>` — most-recently-used at the front |
| `_decisionCacheMeta` | key -> `{estimatedBytes, lruIt}` (iterator into the LRU list) |
| `_decisionCachePinnedKeys` | keys never eligible for eviction |
| `_decisionCacheEstimatedBytes` | running sum of estimated bytes |
| `_decisionCacheBudgetBytes` | budget; default `kDecisionCacheBudgetBytes` = 300 MiB |

### Budget accounting

`EstimateDecisionBytes` is a deliberate over-estimate (512 B per decision + 160 B per item node + key/name char bytes + `sizeof(item)`). Exact allocator introspection is avoided; over-estimation keeps memory conservatively bounded.

`TrackDecisionCacheInsertOrUpdateLocked` adds or updates the entry, splices its LRU node to the front, and adjusts `_decisionCacheEstimatedBytes` by the byte delta. `TrackDecisionCacheEraseLocked` removes the node and subtracts its bytes. `TouchDecisionCacheKeyLocked` (called on a cache hit in `TryGetCachedDecision`) just moves the node to the front.

### Pinning

`SetPinnedFolders(leftRel, rightRel)` rebuilds `_decisionCachePinnedKeys` from the two relative folders and *all their ancestor chains*, always including the root key `.`. The window calls it with the currently visible folders during navigation/refresh (`OnPanePathChanged` path-sync), and with empty paths at run start. Pinned keys protect UI-critical state — the folders the user is looking at and the chain back to the root — from background eviction.

### Eviction

`MaybeEvictDecisionCacheLocked` runs after inserts (`GetOrComputeDecision`, `ScanWorker`, `ApplyPendingContentCompareUpdatesLocked`). It is a no-op while `_decisionCacheEstimatedBytes <= _decisionCacheBudgetBytes`. Otherwise it walks the LRU from the tail (least-recently-used) and evicts candidates, but **skips**:

- the root key `.`,
- pinned keys,
- keys with a queued subtree update (`_pendingSubdirUpdates`),
- decisions whose `anyPending` is true (eviction would lose in-flight progress).

Eviction is itself bounded per call (`kMaxEvictionsPerCall` = 64, `kMaxKeysInspectedPerCall` = 16384) so it never holds `_mutex` long enough to stall other threads — including the UI thread. Evicting a key also drops its `_pendingContentCompareUpdates` entry.

> `SetDecisionCacheBudgetBytesForSelfTest` (only under `ENABLE_TESTS`) lets the compare self-tests shrink the budget to a few KiB to exercise eviction without allocating hundreds of MB. No production code relies on it.

## The cache-only filesystem wrapper

`CompareDirectoriesFileSystem::ReadDirectoryInfo` is the enumeration seam and is strictly cache-only:

1. If compare is disabled (`IsCompareEnabled()` false) or the path is outside the pane's root (`TryMakeRelative` returns `nullopt`), it delegates to the base filesystem so the pane browses normally.
2. Otherwise it calls `session->TryGetCachedDecision(rel)` — **no synchronous I/O, never triggers traversal**.
3. On a miss it calls `session->RequestScanForFolder(rel)` (queues a high-priority background scan) and returns an *empty* enumeration.
4. On a hit it projects the decision to one pane: it skips items absent on this side, skips identical items unless `showIdenticalItems` (pending items are always included), and builds a `FileInfo` buffer via `BuildFilesInformation`.

All mutation operations (`CopyItem(s)`, `MoveItem(s)`, `DeleteItem(s)`, `RenameItem(s)`) delegate straight to the base filesystem; the wrapper never filters them (sync-copy filtering happens at the window/diff-manifest level — see [FileOperations.md](../FileOperations.md)).

## Posted-message protocol

The engine never touches the UI directly. It exposes three callbacks, each stored as an `atomic<shared_ptr<const ...>>` so they can be swapped safely from any thread:

| Callback | Setter | Engine notifier |
|----------|--------|-----------------|
| `ScanProgressCallback` | `SetScanProgressCallback` | `NotifyScanProgress` |
| `ContentProgressCallback` | `SetContentProgressCallback` | `NotifyContentProgress` |
| `DecisionUpdatedCallback` | `SetDecisionUpdatedCallback` | `NotifyDecisionUpdated` |

`SetSessionCallbacksForRun(runId)` (in `CompareDirectoriesWindow.Progress.cpp`) installs callbacks that capture the window `HWND` and the current `runId`, then `PostMessage` a payload to the window. The message IDs are reserved in `Common/WindowMessages.h` (no `WM_APP`/`WM_USER` IDs are defined outside that file):

| Message | Value | Carries |
|---------|-------|---------|
| `kCompareDirectoriesDeferredStart` | `WM_APP + 0x520` | run id in `wParam` (two-phase start) |
| `kCompareDirectoriesScanProgress` | `WM_APP + 0x521` | `ScanProgressPayload` (via `PostMessagePayload`) |
| `kCompareDirectoriesExecuteCommand` | `WM_APP + 0x522` | command id payload |
| `kCompareDirectoriesDecisionUpdated` | `WM_APP + 0x523` | run id in `wParam` |
| `kCompareDirectoriesContentProgress` | `WM_APP + 0x524` | `ContentProgressPayload` |

Payload-bearing messages use `PostMessagePayload(...)` on the engine side and `TakeMessagePayload<T>(lParam)` on the receiver side (helpers in `Common/Helpers.h`). The window drains any undelivered payloads in `WM_NCDESTROY` via `DrainPostedPayloadsForWindow`, so nothing leaks if the window closes mid-run.

### Run-id gating

Every progress message is tagged with the `runId` (the per-window `_compareRunId`, incremented in `PrepareCompareRun`). Handlers ignore any payload whose `runId != _compareRunId` (e.g. `OnScanProgress` early-returns on mismatch, and `kCompareDirectoriesDecisionUpdated` checks `wp == _compareRunId`). This drops late updates from a superseded run.

### Notification throttling and coalescing

The engine throttles `NotifyScanProgress` (80 ms), `NotifyContentProgress` (per-file 80 ms), and `NotifyDecisionUpdated` (120 ms) using `_*LastNotifyTickMs` atomics with a CAS so only one thread wins each window. A **forced** notification is always sent when a scan goes idle or the last content compare completes, so the UI reaches its final state.

On the receiving side the window coalesces further:

- `OnScanProgress` peeks the message queue and drains any newer queued scan-progress messages, keeping only the latest payload (and emits `compare.ui.scan_progress_*` perf counters).
- `kCompareDirectoriesDecisionUpdated` does not refresh inline; it calls `ScheduleDecisionRefresh()`, which arms a single `kCompareDecisionRefreshTimerId` timer (200 ms). `OnDecisionRefreshTimer` then drains updates in bounded batches (`kMaxFlushContentFoldersPerTick` = 8, `kMaxFlushSubdirKeysPerTick` = 16) via the budgeted flush methods, refreshes the panes, and re-arms itself if more work remains.

This keeps the UI thread bounded during large content-compare bursts. The UI thread must never trigger subtree traversal or synchronous compare `ReadDirectoryInfo` — it only reads cached decisions and applies budgeted pending updates.

## Lifecycle and threading rules

- **Start is two-phase.** The window does `PrepareCompareRun` (bump `_compareRunId`, wire callbacks, set roots/settings, `StartScan`) then posts `kCompareDirectoriesDeferredStart` to run `ExecutePreparedCompareRun`. This ensures a Rescan *replaces* run state rather than accumulating it.
- **All session mutation is under `_mutex`**, except the atomics (`_version`, `_backgroundWorkCancelToken`, `_compareEnabled`, `_backgroundWorkEnabled`, the progress counters, and the `atomic<shared_ptr>` callbacks).
- **Deferred cleanup**: large state teardown on Rescan/invalidate/disable is swapped into a `ResetCleanup` and freed on a threadpool thread (`QueueCompareCleanup`), keeping cancel responsive.
- **Worker shutdown** happens in the destructor: `request_stop()` on every `jthread`, clear queues under lock, then `notify_all()` all CVs.

## Extending the engine

- **A new diff criterion** touches `CompareDirectoriesDiffBit`, the comparison logic in `ComputeDecisionForFolder` (and the mirror logic in `ApplyPendingContentCompareUpdatesLocked` for content-driven transitions), and `BuildDetailsTextForCompareItem` in the window for the details line + an `IDS_COMPARE_DETAILS_*` string.
- **Content compare requires `IFileSystemIO` on both sides.** Guard new content paths with `IsContentCompareSupported()` and force `compareContent` off in the options UI when it is unavailable.
- **Cache-key invariants**: keys are produced only by `MakeCacheKey` (generic `/`, `.` for root). New eviction or invalidation code must go through `TrackDecisionCache*Locked` so the LRU and byte total stay consistent.

## Testing

Engine self-tests live in `RedSalamander/SelfTest/CompareDirectories/` and run via `--compare-selftest`. They cover eviction (using `SetDecisionCacheBudgetBytesForSelfTest`), pinning, runtime/remote scenarios, and content-compare transitions. Per the performance contract in `Specs/Core/Core_CompareDirectories.md`, window-facing changes should preserve/extend the `compare.ui.*` metric family, and engine-facing changes should keep deterministic compare self-test coverage alongside any instrumentation.

## See also

- [CompareDirectories.md](../CompareDirectories.md) — user-facing guide.
- [DeveloperGuide.md](../DeveloperGuide.md) — Compare Directories Engine & Window summary in context.
- [FileOperations.md](../FileOperations.md) — sync-copy / diff-manifest behavior driven from compare mode.
- `Specs/Core/Core_CompareDirectories.md` — the normative specification.
