# Filesystem Plugin Improvement Plan (Cross-Plugin Mutation Propagation)

Last updated: 2026-03-07

Status: Implemented in the `codex/plugin-mutation-propagation` worktree.

## Objective

Make visible pane state correct across all built-in file system plugins.

Directory watching is an invalidation source, not the source of truth. The host now owns the refresh, relocation, and mount-retarget decisions needed to keep visible panes accurate when a mutation happens in another pane, in another window, or inside a mounted path.

## Mandatory Behavioral Rules

- If the direct child list of a visible folder changes, the pane refreshes automatically.
- If the current visible folder is deleted, renamed away, or moved outside the current subtree, the pane relocates to the nearest surviving ancestor in the same logical context.
- If no surviving ancestor exists inside that context, the pane falls back to the plugin root.
- If a mounted plugin instance depends on a local backing path and that backing path is renamed or moved, the mounted pane retargets to the new backing path and keeps the internal plugin path when possible.
- If the mounted backing item is deleted or no valid retarget exists, the pane exits the mount and navigates to the nearest surviving `file:` ancestor or the default local root.
- This behavior is mandatory for all built-in plugins used by `FolderView`.
- Off-screen folders may be marked dirty without eager UI refresh.

## Final Design

### 1. Logical filesystem contexts

`DirectoryInfoCache` no longer keys cache and watch state by raw `IFileSystem*`.

The host now identifies a logical context by:

- `pluginId`
- normalized `instanceContext`
- normalized folder path

Multiple live COM instances that represent the same plugin and mount context share cache entries, watchers, and mutation impacts. Raw `IFileSystem*` is now only a live provider handle used to enumerate or watch.

### 2. Central mutation-impact router

All successful host-side mutations and all watch notifications flow through one router in `DirectoryInfoCache`.

The router produces four impact kinds:

- `RefreshCurrentFolder`
- `RelocateCurrentFolder`
- `RetargetInstanceContext`
- `ExitInstanceContext`

This replaced the previous exact-folder dirty ping plus ad hoc `ForceRefresh()` fallback behavior.

### 3. Posted impact delivery

Visible panes are pinned with `WndMsg::kFolderViewDirectoryImpact`.

The host posts `DirectoryImpact` payloads with `PostMessagePayload(...)`, consumes them with `TakeMessagePayload<T>(lParam)`, and drains them on `WM_NCDESTROY`.

`FolderView` still owns refresh debouncing. `FolderWindow` owns the pane-level navigation decisions:

- relocate current folder
- retarget mounted instance context
- exit broken mounts

### 4. Watch model

`IFileSystemDirectoryWatch` remains the single plugin-side invalidation channel.

The host now treats all watcher callbacks as one of two sources:

- real/plugin-native watch events
- synthetic plugin-generated watch events after successful in-app mutations

`overflow = TRUE` always means the incremental event list is not trustworthy and the host must resync the watched folder.

The host consumes each watch notification batch in callback order. It intentionally does not insert a second host-side async reorder stage after callback delivery, because rename routing depends on preserving `RENAMED_OLD_NAME` and `RENAMED_NEW_NAME` adjacency inside a batch.

### 5. Backing-path dependency tracking

Mounted contexts whose `instanceContext` resolves to a Windows path are tracked against that backing path.

This is how read-only `7z` mounts now stay coherent:

- rename or move of the archive retargets the mount
- delete of the archive exits the mount

### 6. Off-screen policy

Pinned visible folders are mandatory watches.

Additional MRU cache entries may still be watched best-effort, but only when the host explicitly opts those entries into off-screen watch coverage.

The current reference implementation does not opt plain cache borrows into that path. By default it keeps:

- cached directory snapshots for performance
- watcher coverage for visible pinned folders
- no automatic watcher attachment for off-screen borrows

Off-screen changes only need to invalidate cache state. They do not need to trigger immediate UI work until the folder becomes visible again.

## Built-in Plugin Matrix

| Plugin | Mutations | Watch mode | External/backend change coverage in v1 | Dependency behavior | Status |
| --- | --- | --- | --- | --- | --- |
| `file` | Yes | Real Win32 watch | Local external changes plus in-app mutations | Source of backing-path invalidation for mounted contexts | Implemented and self-tested |
| `dummy` | Yes | Native plugin watch | Plugin-generated external-style changes plus in-app mutations | None | Implemented and self-tested |
| `7z` | Read-only | None | No direct watch; relies on host dependency routing | Backing archive rename retargets, delete exits mount | Implemented and self-tested |
| `ftp` | Yes | Synthetic `IFileSystemDirectoryWatch` | Successful RedSalamander/plugin-initiated mutations affecting watched folders | None | Implemented; offline contract self-tested |
| `sftp` | Yes | Synthetic `IFileSystemDirectoryWatch` | Successful RedSalamander/plugin-initiated mutations affecting watched folders | None | Implemented; offline contract self-tested |
| `scp` | Yes | Synthetic `IFileSystemDirectoryWatch` | Successful RedSalamander/plugin-initiated mutations affecting watched folders | None | Implemented; offline contract self-tested |
| `imap` | Yes | Synthetic `IFileSystemDirectoryWatch` | Successful RedSalamander/plugin-initiated mutations affecting watched folders | None | Implemented; offline contract self-tested |
| `s3` | Yes | Synthetic `IFileSystemDirectoryWatch` | Successful RedSalamander/plugin-initiated mutations affecting watched folders | None | Implemented; offline contract self-tested |

## Implementation Map

### Host

- `RedSalamander/DirectoryInfoCache.*`
  - logical-context keys
  - provider registration and refcounted context lifetime
  - central mutation router
  - reverse index from Windows backing paths to mounted contexts
- `RedSalamander/FolderWatcher.*`
  - structured relative-path change delivery
  - overflow-to-resync handling
- `RedSalamander/FolderView.*`
  - posted impact handling
  - refresh debouncing via coalescing timer
  - callback seam into `FolderWindow`
- `RedSalamander/FolderWindow.*`
  - provider register/unregister on plugin transitions
  - relocation, retarget, and mount-exit handling
  - focus preservation across refresh/relocate
- `Common/WindowMessages.h`
  - `WndMsg::kFolderViewDirectoryImpact`

### Plugins

- `Plugins/FileSystem`
  - remains the real local watch source
- `Plugins/FileSystemDummy`
  - remains a plugin-native watch source used for deterministic testing
- `Plugins/FileSystemCurl`
  - now exposes `IFileSystemDirectoryWatch` for `ftp`, `sftp`, `scp`, and `imap`
  - emits synthetic notifications after successful create, write, copy, move, delete, and rename operations
- `Plugins/FileSystemS3`
  - now exposes `IFileSystemDirectoryWatch`
  - emits synthetic notifications after successful create, write, and delete operations
- `Plugins/FileSystem7z`
  - still read-only and without direct watch support
  - participates through host backing-path dependency invalidation

## Test Coverage

Build gates run in the independent worktree:

- `.\build.ps1`
- `.\.build\x64\Debug\RedSalamander.exe --fileops-selftest`

Latest verified run: 2026-03-07

- `51` passed
- `0` failed
- `1` skipped

Key new self-test phases:

- `Phase7_CacheBorrowNoWatchInvalidation`
- `Phase7_CrossPaneVisibleRefreshLocal`
- `Phase7_CrossPaneVisibleRefreshDummy`
- `Phase7_CrossPaneRelocateLocal`
- `Phase15_FileSystem7zMountPathImpact`
- `Phase16_RemoteWatchContractExposure`

The skipped case was `Phase16_RemoteS3Sandbox`, which still requires an explicit bucket plus dedicated self-test prefix.

## v1 Limits

- Synthetic remote watches do not poll for unrelated backend-side changes.
- Off-screen folders may remain dirty until they become visible or are borrowed again.
- Plain cache borrows are intentionally not watched unless a future host path opts them in explicitly.
- `7z` remains read-only and does not expose a direct watch interface.
- Remote live integration remains gated by configured connection profiles and sandbox roots.

## References

- `Specs/Core/Core_DirectoryInfoCache.md`
- `Specs/UI/UI_FolderWindow.md`
- `Specs/Plugins/Plugins_VirtualFileSystem.md`
- `RedSalamander/FolderWindow.FileOperations.SelfTest.cpp`
