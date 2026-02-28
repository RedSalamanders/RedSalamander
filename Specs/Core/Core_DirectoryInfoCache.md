# DirectoryInfoCache Specification

## Overview

`DirectoryInfoCache` is an **app-global** (process-wide) cache that stores directory enumeration snapshots (`IFilesInformation`) returned by the `IFileSystem` plugin.

Goals:
- Share `IFilesInformation` results across **all views** (FolderView / NavigationView, and all windows).
- Avoid redundant `IFileSystem::ReadDirectoryInfo()` calls when the same folder is requested multiple times.
- Bound memory usage via **LRU eviction by bytes**, using `IFilesInformation::GetAllocatedSize()` as the per-entry weight.
- Integrate change notifications via `FolderWatcher` to mark cached entries as **dirty** and notify visible views.

Non-goals (v1):
- No incremental patching/diffing of directory snapshots (refresh is full re-enumeration).
- No persistent cache across process restarts.

## Key Concepts

### Cache Key

Each cache entry is identified by:
- `IFileSystem*` pointer identity (so different plugins do not share entries)
- a **normalized absolute folder path** (Windows-style `\` separators, no trailing `\` except drive roots)

Path comparisons are case-insensitive (Windows semantics).

### Snapshot Weight (bytes)

Each cached entry has a byte weight:
- `entryBytes = IFilesInformation::GetAllocatedSize()`

This intentionally tracks the buffer capacity owned by the plugin result object (not `GetBufferSize()`), because it is the value that directly impacts memory pressure.

### LRU-by-bytes eviction

The cache maintains an MRU→LRU list.

When inserting or refreshing an entry causes `currentBytes > maxBytes`, the cache evicts from the LRU end until under the limit.

**Pinned** or **borrowed** entries are not evicted:
- **Pinned**: folder currently displayed in any FolderView (`pinCount > 0`)
- **Borrowed**: temporary read access (`borrowCount > 0`)

If the cache is over budget but cannot evict enough due to pinned/borrowed entries, it stays over budget (and logs telemetry).

## API Shape (host-side)

### Borrowed access (short-lived)

- Used for background enumeration (FolderView worker thread)
- Used for synchronous consumers that do not want to hold a long-lived reference

`BorrowDirectoryInfo(fileSystem, folder, mode)`:
- `BorrowMode::AllowEnumerate`: cache miss/dirty triggers `ReadDirectoryInfo()`
- `BorrowMode::CacheOnly`: never enumerates; returns `S_OK` only if a snapshot exists

The returned `Borrowed` handle:
- holds `borrowCount` while alive
- guarantees the underlying snapshot cannot be evicted during that time

### Pinned folders (on-screen)

FolderView pins the currently displayed folder:
- `PinFolder(fileSystem, folder, hwnd, message)`

Pinning:
- increments `pinCount`
- registers `(hwnd, message)` as a subscriber for “folder dirty” notifications
- enables watchers for pinned folders (subject to watcher limits)

Unpinning:
- decrements `pinCount`
- removes the subscriber

### Clearing entries for a plugin (v2)

When switching/unloading file system plugins, the host may purge cached snapshots for a specific `IFileSystem*`:
- `ClearForFileSystem(fileSystem)`

This drops cached `IFilesInformation` instances (and stops any watchers) for the specified plugin, so the DLL can be safely unloaded.

## Change Notifications (FolderWatcher + plugin watch)

### Watch selection policy

Watchers are created for:
- all pinned entries first (most important: visible folders)
- plus up to `mruWatched` additional MRU entries (best-effort),
  while respecting `maxWatchers`

### Preferred mechanism

`DirectoryInfoCache` uses `IFileSystemDirectoryWatch` (queried from the active `IFileSystem` instance) to watch folders, including virtual file systems like `FileSystemDummy`.

This includes the built-in local file system plugin (`builtin/file-system`), which implements `IFileSystemDirectoryWatch` internally using Win32 directory change notifications.

If the interface is not present (or `WatchDirectory` fails), the folder is treated as **not watched** and the UI uses explicit `ForceRefresh()` fallbacks after mutations.

### Dirty marking

When a watcher receives any change notification:
- mark the corresponding cache entry as `dirty`
- post `message` to all subscribers via `PostMessageW(hwnd, message, 0, 0)`
- coalesce notifications: only one “dirty” message is posted per dirty state (`notifyPosted`)

Consumers decide when to refresh (FolderView schedules a background refresh with debouncing).

## FolderView Integration

- FolderView pins the currently displayed folder via `DirectoryInfoCache::PinFolder(...)`.
- FolderView enumerations use `BorrowMode::AllowEnumerate` to fetch the snapshot:
  - cache hit: reuse existing snapshot
  - cache miss/dirty: perform `ReadDirectoryInfo()` once and share the result with other views
- FolderView receives `WndMsg::kFolderViewDirectoryCacheDirty` (`WM_APP + 0x304`) and schedules a refresh without clearing current items (debounced).
- When a folder is actively watched (`DirectoryInfoCache::IsFolderWatched(...) == true`), the UI relies on watch notifications to refresh after mutations; otherwise it uses explicit `ForceRefresh()` as a fallback.

## NavigationView Integration (siblings dropdown)

NavigationView uses the same cache to collect sibling folders:
- Tries `BorrowMode::CacheOnly` on the parent folder snapshot.
- If unavailable, it schedules a background prefetch (plugin enumeration) to warm the cache and avoids blocking the UI thread.

## Plugin Buffer Trimming (memory feasibility)

Caching directory listings across many folders is only viable if the plugin result buffers are not excessively over-allocated.

The reference plugin (`FileSystem.dll`) applies a **trim heuristic** after enumeration:
- It may reduce its backing buffer to the used byte count when waste is meaningful.
- It skips trimming when the expected reallocation/copy cost is not worth the memory saved.

This directly improves cache effectiveness because `GetAllocatedSize()` becomes a closer approximation of real per-folder memory usage.

## Settings

Settings are stored in the main settings file (`RedSalamander-<Major>.<Minor>.settings.json` in Release, `RedSalamander-debug.settings.json` in Debug) under:
- `cache.directoryInfo`

Keys:
- `maxBytes` (optional, integer|string): maximum total bytes in cache (integer is KiB; string supports `"512MB"`, `"7GB"`; units are base 1024, case-insensitive).
- `maxWatchers` (optional, integer): maximum active watchers.
- `mruWatched` (optional, integer): additional MRU watched folders (best-effort).

Defaults (when missing):
- `maxBytes`: ~6.25% of physical RAM, clamped to `[256 MiB, 4 GiB]`
- `maxWatchers`: `64`
- `mruWatched`: `16`

Hard caps (implementation safety limits):
- `maxWatchers ≤ 1024`
- `mruWatched ≤ 256`

## Telemetry

The cache logs (Debug channel):
- configured limits at startup
- evictions (path + bytes freed)
- enumeration failures

`DirectoryInfoCache::GetStats()` provides programmatic access to current counters and sizes.
