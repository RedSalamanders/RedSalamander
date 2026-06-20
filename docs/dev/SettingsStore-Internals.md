# Settings Store Internals

This page documents the runtime mechanics of RedSalamander's Settings Store for maintainers: how the file watcher hot-reloads external edits, what is required to bump `schemaVersion`, and the default-drop invariant that `PrepareForSave` enforces before every write. For the on-disk format and manual-edit workflow see [../SettingsFile.md](../SettingsFile.md); for the broader subsystem tour see the "Settings & SettingsStore" section of [../DeveloperGuide.md](../DeveloperGuide.md).

## Key files

| File / type | Role |
|---|---|
| `Common/SettingsStore.h` (`Common::Settings::Settings`) | Canonical typed model; `schemaVersion = 16` |
| `Common/Common/SettingsStore.cpp` | Load/save engine, path resolution, atomic write, recovery, file stamp |
| `RedSalamander/SettingsHotReload.{h,cpp}` | Directory watcher thread, save+schema wrapper, stamp dedupe, runtime merge, conflict prompts |
| `RedSalamander/SettingsSave.h` (`SettingsSave::PrepareForSave`) | Drops default-valued optional sections before writing |
| `RedSalamander/SettingsSchemaExport.{h,cpp}` | Aggregated JSON Schema = base + per-plugin config `$defs` |
| `RedSalamander/SettingsSchemaParser.{h,cpp}` | Parses `x-ui-*` schema annotations to drive Preferences panes |
| `Common/WindowMessages.h` | `kSettingsFileChanged`, `kSettingsReloadedFromDisk` window messages |

## Hot-reload control flow

Hot reload keeps the running app, the on-disk file, and any open Preferences dialog in sync after an external edit (the user editing the JSON5 by hand, RedConfigure, or another window saving). The watcher runs off the UI thread but **only posts messages** — every mutation of `g_settings` happens on the UI thread.

### Watcher thread

`SettingsHotReload::Start(targetWindow, appId)` (in `SettingsHotReload.cpp`) first calls `Stop()`, validates the window and app id, then resolves the settings directory from `Common::Settings::GetSettingsPath(appId).parent_path()`. It records an initial `SettingsFileStamp` as `lastAppliedStamp`, creates a manual-reset stop event, and launches a `std::jthread`:

```cpp
g_state.watchThread = std::jthread([targetWindow, stopHandle, settingsDirectory](std::stop_token) noexcept
{ WatchSettingsDirectoryThread(targetWindow, stopHandle, settingsDirectory); });
```

`WatchSettingsDirectoryThread` calls `FindFirstChangeNotificationW` on the directory (not the file) with the filter `FILE_NOTIFY_CHANGE_FILE_NAME | FILE_NOTIFY_CHANGE_LAST_WRITE | FILE_NOTIFY_CHANGE_SIZE`, then loops on `WaitForMultipleObjects` over two handles: the stop event and the change-notification handle.

- Stop event signaled (`WAIT_OBJECT_0`) -> return and end the thread.
- Change signaled (`WAIT_OBJECT_0 + 1`) -> post `WndMsg::kSettingsFileChanged` to the target window with a `SettingsFileChangedPayload` (carrying `GetTickCount64()`) via `PostMessagePayload`, then re-arm with `FindNextChangeNotification`. If re-arm fails, the handle is reset and re-created on the next loop iteration.
- `WAIT_FAILED` -> log and return.

If `FindFirstChangeNotificationW` fails it logs via `Debug::ErrorWithLastError`, waits 1 second on the stop event, and retries (so a transiently missing directory does not kill the watcher).

`Stop()` sets the stop event, `join()`s the thread, and clears all `g_state` fields (including the participant set and both stamps). All `g_state` access is guarded by a single `std::mutex`.

### UI-thread handler

`WndMsg::kSettingsFileChanged` routes to `OnMainWindowSettingsFileChanged` (in `RedSalamander.cpp`), which calls `SettingsHotReload::TryLoadChangedSettings()` and switches on the returned `ChangedSettingsStatus`:

| Status | Handler action |
|---|---|
| `NoChange` | Stamp matched a remembered stamp (self-write or already-rejected) -> return, no reload |
| `Missing` | File absent (`S_FALSE`) -> return |
| `Invalid` | `ERROR_INVALID_DATA` -> `MarkRejectedStamp`, `ShowInvalidReloadAlert` (non-destructive; keeps current settings) |
| `Error` | Other failure HRESULT -> log a warning, return |
| `Loaded` | Proceed to merge and apply |

On `Loaded`, the handler snapshots the live runtime state into `runtimeSettings` (via `CaptureRuntimeSettings`), then:

```cpp
g_settings = SettingsHotReload::MergeDiskSettingsWithRuntimeSession(loadResult.settings, runtimeSettings, CollectRuntimeSettingsWindowIds());
ShortcutDefaults::EnsureShortcutsInitialized(g_settings);
SettingsHotReload::ClearInvalidReloadAlert();
ApplyCurrentSettingsToRunningApp(hWnd);
RefreshRunningPluginsFromSettings(hWnd);
// then MarkAppliedStamp(loadResult.stamp) and NotifyParticipants()
```

`NotifyParticipants()` posts `WndMsg::kSettingsReloadedFromDisk` to every registered participant window. Dialogs that mirror settings (the Preferences dialog calls `SettingsHotReload::RegisterParticipant(dlg)` on init) handle that message to reconcile their own working copy against disk.

### SettingsFileStamp dedupe

The store does not diff file contents to decide whether a change is "real". Instead `Common::Settings::SettingsFileStamp` is a cheap identity built from `BY_HANDLE_FILE_INFORMATION` in `TryGetSettingsFileStampForPath` (opened with `FILE_READ_ATTRIBUTES` and full sharing):

| Field | Source |
|---|---|
| `volumeSerialNumber` | `dwVolumeSerialNumber` |
| `fileIndexHigh` / `fileIndexLow` | `nFileIndexHigh` / `nFileIndexLow` |
| `lastWriteTime` | `ftLastWriteTime` (combined to a `uint64_t`) |
| `fileSize` | `nFileSizeHigh` / `nFileSizeLow` |

The struct has a defaulted `operator==`. `TryLoadChangedSettings` reads the current stamp first; if `ShouldIgnoreStampLocked` finds it equal to either `lastAppliedStamp` **or** `lastRejectedStamp`, it returns `NoChange` without loading. This is what suppresses the watcher echoing the app's own writes and prevents an already-rejected invalid file from re-alerting on every spurious directory notification.

Stamp bookkeeping:

- `MarkAppliedStamp(stamp)` records the stamp and clears `lastRejectedStamp`. Called after a successful save (inside `SavePreparedSettingsAndSchema`) and after a successful reload.
- `MarkRejectedStamp(stamp)` records a stamp that failed to parse, so the bad file is not re-processed until it changes again.

Because saving and reloading both update `lastAppliedStamp`, **always save through `SettingsHotReload::SaveSettingsAndSchema`** (which calls `SettingsSave::PrepareForSave`, then `Common::Settings::SaveSettings`, then refreshes the applied stamp). Writing the file by another path leaves the watcher unable to distinguish your write from an external edit.

When loading does run, `TryLoadChangedSettings` uses `TryLoadSettingsNoRecovery` (which never backs up or falls back to defaults) and retries up to 6 times with a 25 ms sleep on retryable I/O errors (`ERROR_SHARING_VIOLATION`, `ERROR_LOCK_VIOLATION`, `ERROR_ACCESS_DENIED`) to ride out a concurrent writer's atomic replace.

### MergeDiskSettingsWithRuntimeSession

A hot reload must not clobber volatile session state that the user never edits in the file (current window placement, active pane, pane folders). `MergeDiskSettingsWithRuntimeSession(diskSettings, runtimeSettings, runtimeWindowIds)` starts from the freshly-loaded disk settings and overlays the live runtime values:

- For each id in `runtimeWindowIds`, copies that window's placement from `runtimeSettings.windows` over the merged result.
- If runtime `folders` are present, overlays `active`, `layout`, `history`, and `historyFilters`, and for each runtime `FolderPane` updates the matching slot's `current` (or appends a new pane for a slot not on disk).

Everything else (theme, plugins, file actions, etc.) comes from disk, so the external edit wins for those sections while the user's live navigation state survives the reload.

## Bumping `schemaVersion`

The store performs **no field-level upgrade**. The current version is `16`, declared in three coupled places that must all agree:

| Location | Code |
|---|---|
| Model default | `Common/SettingsStore.h`: `uint32_t schemaVersion = 16;` in `struct Settings` |
| Load check | `Common/Common/SettingsStore.cpp` (`LoadSettingsFromResolvedPath`): `if (schemaVersion != 16)` -> `UnsupportedSchemaVersion`; and the trailing `out.schemaVersion = 16;` |
| Serialize | `Common/Common/SettingsStore.cpp` (`SaveSettings`): `yyjson_mut_obj_add_int(doc, root, "schemaVersion", 16);` |

### Replace-on-mismatch behavior

On load, after the JSON parses to an object, `LoadSettingsFromResolvedPath` reads `schemaVersion`:

1. Missing or non-integer -> `RecoverSettingsLoadFailure` with `MissingSchemaVersion` (`ERROR_INVALID_DATA`).
2. Present but not exactly `16` -> `UnsupportedSchemaVersion`, carrying the offending version in `SettingsLoadRecoveryInfo::unsupportedSchemaVersion`.

What `RecoverSettingsLoadFailure` does next depends on the caller's `fallbackToDefaults` / `backupBadFile` flags:

- Cold start (`LoadSettings` / `LoadSettingsWithRecoveryInfo`, both flags true): the bad file is renamed to `*.bad.<UTC timestamp>` by `BackupBadSettingsFile`, defaults are restored, and the function returns `S_FALSE`. There is no migration — a v15 (or any non-16) file is discarded, not upgraded.
- Hot reload (`TryLoadSettingsNoRecovery`, both flags false): the original failure HRESULT is returned and the file is left untouched. The hot-reload path surfaces this as `ChangedSettingsStatus::Invalid` and shows a modeless alert without mutating settings.

A v16 file is also rejected as `LegacyShape` if it still carries the pre-v16 `viewers` / `editors` root keys or `extensions.openWithViewerByExtension` (viewer/editor routing moved into `fileActions` at v16).

### Checklist for raising the version to N

1. `Common/SettingsStore.h`: change the `schemaVersion = 16` default in `struct Settings` to `N` (and adjust any field types you are introducing).
2. `Common/Common/SettingsStore.cpp`, `SaveSettings`: change `yyjson_mut_obj_add_int(doc, root, "schemaVersion", 16)` to `N`. `SaveSettings` is the single serializer (it builds the whole `yyjson_mut_doc` inline); update the per-section emit code here for any added or reshaped fields.
3. `Common/Common/SettingsStore.cpp`, `LoadSettingsFromResolvedPath`: change `if (schemaVersion != 16)` and the trailing `out.schemaVersion = 16;` to `N`. Add or adjust the relevant `Parse*` helper for the new shape (the load path dispatches one `Parse*` per section: `ParseWindows`, `ParseTheme`, `ParsePlugins`, `ParseExtensions`, `ParseFileActions`, `ParseUserMenuSettings`, `ParseMakeFileListSettings`, `ParseShortcuts`, `ParseCache`, `ParseFolders`, `ParseMonitor`, `ParseMainMenu`, `ParseStartup`, `ParseUi`, `ParseConnections`, `ParseFileOperations`, `ParseCompareDirectories`, `ParseHotPaths`, `ParseSelectionMasks`, `ParseSearchSettings`, `ParseBatchRenameSettings`).
4. If the new field is user-facing, add its `x-ui-*`-annotated entry to the base schema `SettingsStore.schema.json` so `SettingsSchemaParser` surfaces it in a Preferences pane, and (if it is an optional default-droppable section) extend `SettingsSave::PrepareForSave` — see below.
5. Update `Specs/Core/Core_SettingsStore.md` (the data-model section is versioned, e.g. "Settings Data Model (v16)") and the example payloads, plus [../SettingsFile.md](../SettingsFile.md) if the on-disk shape visibly changes.

Because there is no upgrader, bumping the version means every existing user's file becomes `UnsupportedSchemaVersion` and is backed up + reset on next cold start. Treat a bump as a hard reset of persisted state, not a migration.

## The `PrepareForSave` default-drop invariant

`SettingsSave::PrepareForSave(const Settings&)` (header-only in `RedSalamander/SettingsSave.h`) runs on every save, before `Common::Settings::SaveSettings`, and produces a copy with **default-valued optional sections dropped** (`.reset()`-ed to `std::nullopt`). The invariant is: *an optional section is written to disk only if it differs from its constructed default*. This keeps the file small and human-readable and means a user who never touched a feature has no entry for it.

Sections it can drop, and the rule applied:

| Section | Dropped when |
|---|---|
| `shortcuts` | `ShortcutDefaults::AreShortcutsDefault(...)` is true |
| `monitor` | menu flags and the low 6 bits of `filter.mask` and `filter.preset` all equal `MonitorSettings` defaults |
| `cache` | no `directoryInfo` field was written (no positive `maxBytes`, no `maxWatchers`, no `mruWatched`) |
| `fileOperations` | every field equals `FileOperationsSettings` defaults (and all optional/string fields empty) |
| `compareDirectories` | every field equals `CompareDirectoriesSettings` defaults and both ignore-pattern lists empty |
| `hotPaths` | no slot holds a non-empty path and `openPrefsOnAssign` is false |
| `selectionMasks` | all three history lists empty |
| `search` | every field equals `SearchDialogSettings` defaults and all lists/strings empty |
| `batchRename` | every field equals `BatchRenameSettings` defaults and all recent-pattern lists empty |
| `makeFileList` | `value() == MakeFileListSettings{}` |
| `ui` | `value() == UiSettings{}` |

Consequences to keep in mind when changing these structs:

- Adding a field to one of these sections means extending the corresponding `hasNonDefault` test in `PrepareForSave`. If you forget, a user setting only the new field will have the whole section dropped on save and silently lost.
- For sections compared via `operator==` (`makeFileList`, `ui`), adding a member just works once the struct's defaulted equality includes it.
- The drop is non-destructive at load time: a missing section parses back to its default, so round-tripping a dropped section reproduces the same in-memory state.
