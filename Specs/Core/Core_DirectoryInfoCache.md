# DirectoryInfoCache Specification

## Overview

`DirectoryInfoCache` is an app-global cache for folder enumeration snapshots (`IFilesInformation`) returned by `IFileSystem`.

Its job is no longer just cache reuse. It is also the host-side mutation router that keeps visible panes coherent across:

- multiple panes
- multiple windows
- multiple live `IFileSystem` COM instances for the same plugin context
- mounted file systems whose backing path changes in the local file system

## Goals

- Share directory snapshots across views and windows.
- Key cache state by logical plugin context instead of raw COM identity.
- Bound memory by LRU-by-bytes eviction.
- Keep visible folders correct after direct and indirect mutations.
- Use watchers for invalidation, not as the only correctness mechanism.

## Cache Identity

### Logical context key

Each entry is keyed by:

- `pluginId`
- normalized `instanceContext`
- normalized folder path

`IFileSystem*` is only a provider handle. `RegisterProvider(...)` and `UnregisterProvider(...)` maintain the mapping from one or more live COM objects to a logical context.

This means:

- two panes using different COM instances for the same plugin and mount context still share cache and watch state
- a context is cleared only when its last registered provider disappears

### Path normalization

- Paths are normalized before lookup.
- Key comparisons are case-insensitive.
- For mounted plugins, the folder path is the plugin path. The mount identity lives in `instanceContext`.

## Snapshot Weight And Eviction

- `entryBytes = IFilesInformation::GetAllocatedSize()`
- eviction order is MRU to LRU by bytes
- pinned or borrowed entries are not evicted

If the cache cannot get under budget because visible folders are pinned, it remains over budget until pressure drops.

## Public Host API Shape

### Borrowed access

`BorrowDirectoryInfo(fileSystem, folder, mode)`:

- `BorrowMode::AllowEnumerate`
- `BorrowMode::CacheOnly`

Borrowed handles keep the snapshot alive and non-evictable for the duration of use.

Pure cache borrows do not automatically attach a watcher. In the reference implementation, watcher attachment is a visible-pane decision by default.

### Visible pinning

`PinFolder(fileSystem, folder, hwnd, message)`:

- increments `pinCount`
- subscribes the visible pane to directory impacts
- makes the folder eligible for mandatory watch coverage

Pinned panes now receive `WndMsg::kFolderViewDirectoryImpact` payloads, not a raw dirty ping.

### Provider lifecycle

`RegisterProvider(...)`:

- associates a live `IFileSystem*` with a logical context
- allows cache/watch sharing across instances

`UnregisterProvider(...)`:

- removes the live provider
- clears the logical context only if no providers remain

`ClearForFileSystem(...)` remains available for explicit host purges, but the normal lifetime is driven by provider registration.

## Mutation Routing

All mutation sources now flow through the same router:

- watcher callbacks
- host-side create/copy/move/delete/rename completions
- cross-file-system bridge operations
- backing-path dependency hits from local `file:` changes

The router exposes:

- `NotifyFolderContentsChanged(...)`
- `NotifyPathCreated(...)`
- `NotifyPathDeleted(...)`
- `NotifyPathMoved(...)`

## Directory Impacts

The router emits one of four payloads:

- `RefreshCurrentFolder`
- `RelocateCurrentFolder`
- `RetargetInstanceContext`
- `ExitInstanceContext`

Rules:

- if the visible folder's direct children changed, refresh that folder
- if the current folder was deleted or moved away, relocate to the nearest surviving ancestor
- if a mounted backing path was renamed or moved, retarget the instance context
- if a mounted backing path was deleted, exit the mount to local file space

## Watch Policy

### Mandatory watches

All pinned visible folders are watched first.

### Best-effort extra watches

Up to `mruWatched` additional MRU entries may be watched while staying under `maxWatchers`, but only when the host explicitly opts an entry into off-screen watch coverage.

The current reference host does not opt plain cache borrows into that path, so the effective default policy is:

- cache borrows: yes
- visible automatic watches: yes
- off-screen automatic watches: no

Off-screen cached folders are still invalidated by routed mutations and may remain dirty without triggering eager UI work.

### Watch sources

`DirectoryInfoCache` uses `IFileSystemDirectoryWatch` when the plugin exposes it.

That watch source may be:

- real OS-backed notifications
- native plugin notifications
- synthetic plugin-generated notifications after successful in-app mutations

The host consumes each `IFileSystemDirectoryWatch` notification batch in callback order. It intentionally does not add a second host-side async reorder stage after watch delivery, because rename routing depends on preserving `RENAMED_OLD_NAME` and `RENAMED_NEW_NAME` adjacency inside a batch.

If a plugin does not expose a watch interface, the cache still stays correct through explicit mutation routing from host-side operations and, for mounted contexts, backing-path dependency tracking.

### Overflow handling

`overflow = TRUE` means incremental events are not trustworthy. The cache marks the watched folder dirty and posts a refresh impact so the visible pane performs a full re-enumeration.

Malformed rename batches are handled conservatively: an unmatched trailing `RENAMED_OLD_NAME` is treated as delete, while unmatched `RENAMED_NEW_NAME` falls back to parent-folder refresh.

## Mounted Backing-Path Dependencies

When a plugin `instanceContext` resolves to a Windows path, `DirectoryInfoCache` tracks a reverse dependency from that local path to the mounted logical context.

Current use:

- `7z:` mounts

Behavior:

- local rename or move retargets the mounted instance context
- local delete exits the mount

## FolderView And FolderWindow Integration

`FolderView`:

- pins the current folder
- receives `WndMsg::kFolderViewDirectoryImpact`
- debounces refresh impacts locally
- forwards relocate/retarget/exit impacts to `FolderWindow`

`FolderWindow`:

- owns navigation changes caused by impacts
- preserves focus where possible during refresh and relocation
- exits or retargets mounts when backing paths change

All posted impacts use the standard payload helpers:

- `PostMessagePayload(...)`
- `TakeMessagePayload<T>(lParam)`
- `DrainPostedPayloadsForWindow(hwnd)`

## Settings

Settings live under `cache.directoryInfo`:

- `maxBytes`
- `maxWatchers`
- `mruWatched`

Defaults:

- `maxBytes`: about 6.25% of physical RAM, clamped to `[256 MiB, 4 GiB]`
- `maxWatchers`: `64`
- `mruWatched`: `16`

## Telemetry

`DirectoryInfoCache::GetStats()` reports:

- cache size and limits
- hits, misses, enumerations, evictions
- dirty marks
- watcher counts
- pinned entry count

Debug logging covers evictions, enumeration failures, and related cache events.
