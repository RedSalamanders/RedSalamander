# Settings Store Specification (Common.dll)

## Overview

RedSalamander needs a **single, shared** settings layer implemented in **`Common.dll`** to persist user parameters in a Windows-correct location. Settings are stored as **JSON** and must be **read/written with `yyjson`** (vcpkg-managed).

This specification defines:
- Storage location and file naming on Windows
- Read/write guarantees (atomic save, recovery behavior)
- Settings data model (window placement, theme system, multi-pane folder state)
- The JSON Schema used to validate the file format

## Goals

- Provide a **per-user** settings store for RedSalamander components.
- Store settings in **`%LocalAppData%`** using the **Known Folder** API (no hard-coded paths).
- Use **`yyjson`** for parsing and serialization.
- Persist and restore:
  - Window size/position (+ state) with a **full-visibility** restore check
  - Current theme selection and **custom user themes**
  - Multi-pane folder state (current folder + navigation history)
- Support versioned evolution via `schemaVersion`.
- Reader accepts JSON5 (comments, trailing commas) for user-friendly editing.
- Writer emits strict JSON and omits default values; comments/trailing commas are not preserved.


## Non-Goals

- No registry storage for these settings (registry-backed settings can remain where they are until explicitly migrated).
- No runtime JSON Schema validation dependency (the schema is normative; runtime validation is implemented as type/range checks when reading).

## Storage Location (Windows)

### Base directory

Settings are stored under the per-user **Local AppData** directory:
- Known folder: `FOLDERID_LocalAppData`
- Subdirectory: `RedSalamander\\Settings`

Example (typical):
- `C:\\Users\\<User>\\AppData\\Local\\RedSalamander\\Settings\\`

### File naming

To avoid cross-process contention and keep settings scoped correctly, each executable uses its own file:
- Debug builds: `<AppId>-debug.settings.json`
- Release builds: `<AppId>-<Major>.<Minor>.settings.json` (from `Common/Version.h`: `VERSINFO_MAJOR` and `VERSINFO_MINOR`)
- Legacy name (pre-versioning, supported for migration): `<AppId>.settings.json`
- Schema file pattern (always unversioned): `<AppId>.settings.schema.json`
- `AppId` examples:
  - `RedSalamander`
  - `RedSalamanderMonitor`

Debug-load note:
- In Debug builds, if `<AppId>-debug.settings.json` does not exist, settings load falls back to the versioned file (and then the legacy file) so developers can reuse existing settings without copying/renaming files.

Example (Release):
- `C:\\Users\\<User>\\AppData\\Local\\RedSalamander\\Settings\\RedSalamander-7.0.settings.json`

### Schema file

The canonical JSON Schema is stored in the repo at:
- `Specs/SettingsStore.schema.json`

Every explicit `SaveSettingsAndSchema` save writes a schema file next to the settings file:
- `<AppId>.settings.schema.json`

The settings JSON includes a `$schema` property referencing it (relative path):
- `$schema: "./<AppId>.settings.schema.json"`

Notes:
- `Common.dll` writes the base schema (identical to `Specs/SettingsStore.schema.json`) as a best-effort convenience for manual editing.
- `Common.dll` loads the base schema text from `SettingsStore.schema.json` shipped next to the exe (copied from `Specs/SettingsStore.schema.json` during the build) and caches it in memory.
- `RedSalamander.exe` overwrites that file with an aggregated schema that includes plugin configuration schemas under `plugins.configurationByPluginId[pluginId]` (best-effort).
- UI controls that change values without changing schema shape MAY use the process-wide serialized
  asynchronous settings queue. Those saves capture an immutable caller-thread snapshot, debounce and
  coalesce consecutive snapshots for the same app, write only the settings JSON, and leave the
  existing schema file unchanged.
- Confirmed Windows session end submits one bounded, settings-only final request through that same
  serialized coordinator. Entering this final-save mode rejects later submissions, prevents older
  queued snapshots from writing, and serializes the final snapshot behind any write already in
  progress. The handler does not collect plugin schemas or rewrite the schema sidecar.
- Process shutdown stops the directory watcher without waiting for old saves, then closes normal
  submissions and reserves the final-save transition under the same submission lock. The public final-save
  call remains admissible after `BeginProcessShutdown()`, captures one settings/plugin-schema snapshot,
  enqueues exactly one schema-writing request, and waits at most five seconds. Concurrent normal submissions
  fail with `ERROR_SHUTDOWN_IN_PROGRESS`; duplicate finalization reuses the first request's completion and
  never enqueues another snapshot. Timeout returns `ERROR_TIMEOUT`; the request and its completion remain
  worker-owned, and UI teardown continues without a join.
- The serialized worker and hot-reload session state are explicitly process-lifetime in
  `RedSalamander.exe`. Entering process shutdown rejects later submissions and prevents late worker
  completion from posting UI payloads or emitting Debug/perf callbacks. This process-lifetime
  exception does not apply to plugin DLL workers.

### UI entry point

`Preferences -> Advanced` exposes a command link that opens `GetSettingsPath(L"RedSalamander")` with the shell default editor for the current user. `Preferences -> Monitor` exposes a separate command link that opens `GetSettingsPath(L"RedSalamanderMonitor")`. Either command may create the target JSON/settings schema first when the file is missing, but invoking the link MUST NOT mark Preferences dirty. The generated `RedSalamander.settings.schema.json` MUST NOT expose the root `monitor` property; the generated `RedSalamanderMonitor.settings.schema.json` MUST expose it.

## Read/Write Requirements

### Encoding

- Settings file (`GetSettingsPath(appId)`): **UTF-8** (with BOM)
- Schema file (`<AppId>.settings.schema.json`): **UTF-8** (no BOM)
- In-memory strings: UTF-16 (`std::wstring` / `std::filesystem::path`) as per project convention

### Atomic saves

Saving must be atomic and conflict-aware to prevent partial writes and stale-snapshot overwrites:

1. A loaded `Settings` snapshot retains the exact identity of the canonical save target read from the same
   handle as its bytes. A missing target is represented explicitly by no stamp.
2. Serialize the complete document before publication. Malformed UTF-16 is a save failure
   (`ERROR_NO_UNICODE_TRANSLATION`); it is never replaced by an empty JSON string.
3. Stage the bytes through `Common::Files::LocalFileTransaction`, which creates a unique sibling in the same
   directory, uses the bounded zero-progress-safe handle writer, verifies the requested size, and flushes it.
4. Under the per-target cross-process commit lock, compare the current target identity with the snapshot's
   expected identity. A mismatch returns `ERROR_REVISION_MISMATCH` and leaves the target unchanged.
5. Publish the unique sibling using `MoveFileExW` replace/write-through semantics. The identity returned to the
   caller is captured from the finalized sibling held through the move; do not re-stat the destination afterward.
6. A successful mutable save advances the caller's expected identity so the same snapshot can be saved again.

The immutable `SaveSettings(..., const Settings&)` overload is a one-shot compatibility surface for shutdown
and focused test snapshots. A caller that may save the same in-memory object again MUST use the mutable overload.
Schema-file publication uses the same unique-sibling transaction but has no settings-snapshot CAS contract.

### Recovery behavior

For the normal startup / explicit recovery path (`LoadSettings(...)`):
- If the file is missing, unreadable, not valid JSON, has an invalid root/schema marker, or uses an
  unsupported older schema, start with defaults.
- If a whole-document failure existed on disk, rename it to a backup for diagnostics:
  - `<SettingsFileName>.bad.<UTC timestamp>`
- Startup callers that need to explain recovery to the user call `LoadSettingsWithRecoveryInfo(...)`. When a previous settings file is backed up and defaults are restored, the app MUST show a localized warning that includes the original settings path, the backup path, and explains that the current run is using default settings. The warning should tell the user to close the app, compare the backup with the new settings file at the original path, and copy back only the settings they still need.
- `fileActions`, `userMenu`, and `shortcuts` are independent optional sections. A malformed one resets only
  that section, leaves every valid section loaded, records a `SettingsSectionRecoveryInfo` entry, and does
  not move or back up the whole settings file. The malformed section is retained as owned opaque JSON so a
  later canonical save does not destroy information the user may need to repair.
- A schema version greater than the current version is forward data, not a corrupt file. Startup uses
  defaults for the running process but leaves the source bytes and path untouched, records the source
  version, and marks the resulting settings snapshot `ExplicitReplacementRequired`. All automatic,
  explicit-normal, asynchronous, and shutdown saves reject that snapshot with `ERROR_REVISION_MISMATCH`.
  RedSalamander MUST surface an actionable warning that automatic persistence is disabled. A deliberate
  user-approved replacement may clear the block only after the source has been moved successfully to the
  standard timestamped backup path. The startup decision defaults to preservation; replacement requires an
  explicit affirmative choice and keeps the exact newer source bytes in that backup.
- When any whole-document recovery attempts to back up an invalid source but the move fails, the defaults
  snapshot MUST also be marked `ExplicitReplacementRequired`. Automatic, asynchronous, and shutdown saves may
  not overwrite the only remaining recovery artifact. Explicit replacement first moves that source to the
  standard backup path, then saves against the now-missing canonical target.

For the non-destructive hot-reload path (`TryLoadSettingsNoRecovery(...)`):
- Missing file returns `S_FALSE`.
- Invalid JSON / invalid types / unsupported schema version returns a failure `HRESULT`.
- The file is left in place; it is **not** renamed to `.bad.*`.
- The caller decides whether to keep current runtime settings, warn the user, or recover in some other way.

### Tolerant reads, canonical writes

- Unknown top-level keys must be copied into `SettingsPersistenceState::opaqueTopLevelMembers` as owned
  `JsonValue` data and written back unchanged at the JSON-value level (forward compatibility). No yyjson
  pointer or borrowed string may survive the parsed document lifetime.
- Missing keys use defaults.
- Writer emits strict JSON with stable formatting and ordering for present keys, but **omits default values** (and whole sections) when there is nothing meaningful to persist.
- Reader accepts JSON5 features (comments, trailing commas) and UTF-8 BOM; writer outputs strict JSON and does not preserve comments/trailing commas.

## Common.dll API Surface (v2)

The settings store is implemented in `Common.dll` and consumed by executables.

### Responsibilities

`Common.dll` provides:
- Path resolution (`FOLDERID_LocalAppData` + subdirectory + per-build settings file name; see above)
- Load/Save for a single strongly-typed settings object
- Conversion helpers (UTF-16 ⇄ UTF-8, color parsing, window placement normalization)

### Suggested C++ API (shape only)

This is the intended public shape (names can be adjusted during implementation):

```cpp
namespace Common::Settings
{
    struct WindowPlacement;
    struct ThemeDefinition;
    struct Settings;
    struct SettingsFileStamp;

    std::filesystem::path GetSettingsPath(std::wstring_view appId) noexcept;
    std::filesystem::path GetSettingsSchemaPath(std::wstring_view appId) noexcept;

    // Canonical base schema text (UTF-8 JSON, no BOM).
    std::string_view GetSettingsStoreSchemaJsonUtf8() noexcept;

    HRESULT LoadSettings(std::wstring_view appId, Settings& out) noexcept;
    HRESULT LoadSettingsWithRecoveryInfo(std::wstring_view appId, Settings& out, SettingsLoadRecoveryInfo* recovery) noexcept;
    HRESULT TryLoadSettingsNoRecovery(std::wstring_view appId, Settings& out) noexcept;
    HRESULT TryGetSettingsFileStamp(std::wstring_view appId, SettingsFileStamp& out) noexcept;
    HRESULT BackupSettingsForExplicitReplacement(std::wstring_view appId, std::filesystem::path& backupPath) noexcept;
    HRESULT SaveSettings(std::wstring_view appId, Settings& settings) noexcept;
    HRESULT SaveSettings(std::wstring_view appId, const Settings& oneShotSettings) noexcept;
    HRESULT SaveSettingsValuesOnly(std::wstring_view appId, Settings& settings) noexcept;
    HRESULT SaveSettingsValuesOnlyWithStamp(
        std::wstring_view appId, Settings& settings, SettingsFileStamp& writtenStamp) noexcept;

    // Writes `<AppId>.settings.schema.json` next to the settings file.
    HRESULT SaveSettingsSchema(std::wstring_view appId, std::string_view schemaJsonUtf8) noexcept;
}
```

Parsed `JsonValue` consumers use the shared `Common::Settings::FindMember`,
`GetString`, `GetWString`, `GetBool`, `GetUInt32`, and `GetArray` accessors rather
than duplicating variant/object traversal. Typed accessors return no value for a
missing or wrong-typed member; `GetUInt32` also rejects negative and overflowing
numbers, and `GetWString` rejects malformed UTF-8 instead of substituting a
replacement character.

`SettingsLoadRecoveryInfo` reports whether the startup recovery path used defaults, whether an existing file was moved to a backup, the original settings path, the backup path, the recovery reason, the unsupported schema version when that was the failure, and zero or more section-scoped recovery records. `LoadSettings(...)` is the compatibility wrapper for callers that do not need this detail.

### File-stamp helper

`SettingsFileStamp` is a stable identity/value snapshot of the current settings file used both for optimistic
save concurrency and by live-reload callers to de-duplicate notifications and suppress self-saves.

Fields:
- `volumeSerialNumber`
- `fileIndexHigh`
- `fileIndexLow`
- `lastWriteTime`
- `fileSize`

Return contract:
- `TryGetSettingsFileStamp(...)`
  - `S_OK`: stamp retrieved
  - `S_FALSE`: file missing
  - failure `HRESULT`: unexpected I/O/query failure
- `TryLoadSettingsNoRecovery(...)`
  - `S_OK`: settings loaded; malformed recoverable optional sections use their defaults and retain their
    opaque source member; `SettingsPersistenceState::expectedFileStamp` identifies the exact loaded target
  - `S_FALSE`: file missing
  - failure `HRESULT`: invalid/unreadable/unsupported file without fallback or backup
- `SaveSettingsValuesOnlyWithStamp(...)` returns the stamp of the flushed temporary file that was
  atomically moved into place. It MUST NOT re-stat the destination path after replacement, because a
  later external writer may already own that path.
- Save entry points compare `expectedFileStamp` under the cross-process target lock. Both a changed identity and
  an unexpected create/delete are conflicts. `ERROR_REVISION_MISMATCH` means no settings bytes were published.
- Numeric accessors for 32-bit fields MUST reject values greater than `UINT32_MAX`; they MUST NOT truncate.

## Live Reload Semantics (RedSalamander.exe)

`RedSalamander.exe` live-watches only the main settings file returned by `GetSettingsPath(L"RedSalamander")`.

Watcher rules:
- Detection is event-driven (directory change notification), but only the main settings file stamp is authoritative.
- `SettingsHotReload::Start(...)` must not synchronously wait on directory-watch readiness during window creation and must not return `WAIT_TIMEOUT` for a transient initial `FindFirstChangeNotificationW(...)` arming failure.
- Directory watcher readiness is worker-internal state; startup and main-window creation continue after launching the watcher, while the worker retries asynchronously until it can arm.
- The worker creates a missing settings directory before it arms the directory notification, and retry waits remain cancellation-aware so teardown does not wait for the next retry interval.
- The worker captures the initial authoritative settings-file stamp before arming and performs a post-arm catch-up comparison. A replacement made between the initial stamp and successful notification arming MUST post `WndMsg::kSettingsFileChanged` even when no directory event is subsequently delivered.
- After a transient watcher-arm failure, the watcher self-arms on a later retry when the settings directory becomes reachable, and subsequent directory changes still post `WndMsg::kSettingsFileChanged`.
- `RedSalamander.exe` writes its own settings through a shared save helper that opens an internal-save
  epoch, writes values-only, publishes the exact atomic-writer stamp with the originating session
  token, and ends the epoch before aggregate-schema I/O.
- A watcher notification observed while that internal-save epoch is active is deferred and rechecked
  after the exact stamp is published. An external replacement made after the app's atomic move has a
  different stamp and MUST be loaded even when schema generation or another post-write step is still
  running.
- If an internal-save epoch begins or ends between a reload stamp/load check, the reload retries its
  observation. Repeated epoch churn reposts a change notification; it must not silently consume an
  external event, including when the internal save fails.
- The serialized save coordinator may advance queued snapshots only along its own known source-to-commit
  lineage. A snapshot matching neither the original source nor the coordinator's latest commit is an external
  revision and MUST still fail CAS rather than being silently rebased.
- A changed stamp already recorded as `lastAppliedStamp` or `lastRejectedStamp` is ignored on the next reload check.
- `Themes\\*.theme.json5` files are **not** watched in this iteration.
- `RedSalamanderMonitor.exe` settings do not participate in this main-app hot-reload flow.

Preferences commits the main and Monitor settings documents independently because NTFS does not provide an
atomic rename spanning two files. If the main document commits but the Monitor document fails, the main settings
remain committed and are applied to the running app; the dialog advances only the main baseline, keeps Monitor
changes dirty, and shows a localized partial-success error. Because Preferences owns the complete Monitor
section, a Monitor-only revision conflict may load the newest Monitor document, replace that section from the
working copy, and retry CAS once. A second conflict or any other failure remains pending for the user to retry.

Merge policy after a valid external reload:
- Disk is authoritative for persisted main-app user-editable sections such as `theme`, `ui`, `plugins`, `connections`, `extensions`, `shortcuts`, `cache`, `fileOperations`, `compareDirectories`, `hotPaths`, `mainMenu`, `startup`, `search`, and folder preference fields.
- The `monitor` section belongs to the `RedSalamanderMonitor` settings app id. The main `RedSalamander` settings file MUST NOT be used as the persistence owner for Monitor Preferences edits.
- Runtime session state is preserved for already-open windows and current pane navigation.
- Preserved runtime window placements are the placements of currently open modeless/top-level windows already running in the process.
- Preserved folder session fields are:
  - `folders.active`
  - `folders.layout`
  - `folders.history`
  - `folders.historyFilters`
  - `folders.items[*].current`
- External reload must not live-move existing windows or navigate panes to the paths stored on disk.
- Settings-driven plugin rediscovery must capture each live pane's plugin ID, short ID, instance
  context, and raw provider path before releasing providers, qualify that raw path exactly once, then
  restore and re-enumerate both panes after providers reload. A provider refresh must never leave
  retained pane locations with empty item models or double-qualify mounted/archive paths.

Invalid external file behavior:
- Keep the current runtime settings unchanged.
- Return an error from `TryLoadSettingsNoRecovery(...)` to the app-level hot-reload handler.
- Show a single modeless localized warning in the app.
- Do **not** rename/back up the file during the live reload failure path.

Manual association reload:
- `cmd/app/rereadAssociations` uses the same non-destructive `TryLoadSettingsNoRecovery(...)` path so manual reloads never rename or back up an invalid settings file.
- A successful manual reload applies persisted viewer/editor/User Menu action settings, extension associations, plugin settings, shortcuts, and other disk-authored sections while preserving already-open window placement and current pane navigation state.
- After applying the merged settings, the app rebuilds dynamic View With, Edit With, User Menu, ShellNew, and file-system-plugin menus; clears normal and association icon caches before pane refresh can repopulate them; refreshes both panes; preserves the active pane; updates the applied settings stamp when available; and notifies settings-reload participants.
- A failed manual reload keeps the previous runtime settings and reports through the same localized invalid-reload alert used by the external watcher.

## Settings Data Model (v16)

### Root object

The root JSON object may contain (depending on the application):
- `schemaVersion` (integer): format version (current: v16 = `16`); unsupported versions are treated as invalid (file is backed up and defaults are used, no migration).
- Startup recovery for an unsupported older version uses the same `.bad.<UTC timestamp>` backup path and
  user-facing warning as other destructive recoveries. v15 files are intentionally not migrated to v16.
  Unsupported future versions are preserved in place and block persistence as defined under Recovery behavior.
- `windows` (object): per-window placement records
- `theme` (object): current theme + custom themes
- `plugins` (object): plugin discovery + per-plugin configuration
- `ui` (object): app-wide DxUI density, motion, and backdrop preferences
- `mainMenu` (object): RedSalamander main window menu bar state
- `cache` (object): cache configuration (directory enumeration cache, etc.)
- `folders` (object): multi-pane folder state (current folder + global folder history)
- `search` (object): persisted Find Files and Directories dialog state
- `monitor` (object): RedSalamanderMonitor UI state (menu toggles, filter state); persisted by the `RedSalamanderMonitor` app id
- `shortcuts` (object): shortcut key bindings
- `extensions` (object, optional): extension-based behaviors (e.g., open archives as virtual file systems)
- `connections` (object, optional): global Connection Manager settings and saved `ConnectionProfile` entries (non-secret fields only)
- `fileActions` (object, optional): viewer/editor actions and command associations for View, Alternate View, Edit, Alternate Edit, View With, Edit With, and Edit New
- `userMenu` (object, optional): settings-driven external commands for `Commands -> User Menu`, `F9`, and `cmd/pane/userMenu/<itemId>`
- `makeFileList` (object, optional): last selected options for `Commands -> Make File List`

## Plugins (v11)

Plugin settings live under:
- `plugins`

Keys:
- `currentFileSystemPluginId` (string): default `IFileSystem` plugin ID (example: `"builtin/file-system"`)
- `customPluginPaths` (string[]): absolute paths to user-added plugin DLLs
- `disabledPluginIds` (string[]): plugin IDs that must not be loaded on startup
- `configurationByPluginId` (object): per-plugin configuration payloads (JSON object), keyed by plugin ID

Notes:
- Plugin IDs are long, stable identifiers (`builtin/<name>` for embedded/optional, `user/<name>` for custom).
- Each plugin also exposes a unique **short ID** used for navigation prefixes (e.g., `file`, `fk`, `ftp`, `gdrive`, `s3`), but settings use the long ID.
- When migrating older settings, legacy IDs such as `"file"`, `"builtin/filesystem"`, or `"fk"` are normalized to `"builtin/file-system"` and `"builtin/file-system-dummy"`.
- Custom plugins are referenced **in place** (paths in `customPluginPaths`); the host never copies DLLs into the `Plugins` folder.
- Changes made via the **Manage Plugins** UI (add/remove/enable/disable/configure) are saved immediately to reduce the risk of losing configuration on crashes.
- Selecting the active file system plugin (changing `currentFileSystemPluginId`) is saved immediately.
- `configurationByPluginId` stores the plugin's canonical configuration object returned by `IInformations::GetConfiguration()` after applying changes (values may be normalized/clamped by the plugin).
- If a plugin reports `SomethingToSave() == FALSE`, its entry is removed/omitted from `configurationByPluginId`.

### Built-in plugin configuration

`plugins.configurationByPluginId[...]` entries are plugin-defined and intentionally treated as an opaque JSON object by the settings store schema.

Built-in plugin configuration keys are documented in their respective plugin specs (or in the plugin-type spec when appropriate):
- `builtin/file-system-dummy`: `Specs/Plugins/Plugins_VirtualFileSystem.md`
- `builtin/file-system-gdrive`: `Specs/FileSystem/FileSystem_GoogleDrive.md`
- `builtin/viewer-imgraw`: `Specs/Plugins/Plugins_ViewerImgRaw.md`
- `builtin/viewer-space`: `Specs/Plugins/Plugins_ViewerSpace.md`
- `builtin/viewer-text`: `Specs/Plugins/Plugins_ViewerText.md`
- `builtin/viewer-vlc`: `Specs/Plugins/Plugins_ViewerPlugins.md`
- `builtin/viewer-web`, `builtin/viewer-json`, `builtin/viewer-markdown`: `Specs/Plugins/Plugins_ViewerWeb.md`

Current built-in ViewerText configuration includes text/hex defaults plus diff defaults (`diffDefaultLayout`, `diffContextMode`, `diffAutoOpenMode`) persisted under `plugins.configurationByPluginId["builtin/viewer-text"]`.

Current built-in ViewerVLC configuration includes `lastVolumePercent`, `muted`, and `audioVisualization` under `plugins.configurationByPluginId["builtin/viewer-vlc"]`. Fresh ViewerVLC configurations default `audioVisualization` to `visual`; existing persisted values, including `goom`, continue to round-trip. Standalone viewer closes, embedded preview closes, preview-plugin replacement, and app shutdown all persist changed ViewerVLC configuration through the same plugin configuration map.

## Extensions (v11)

Extension settings live under:
- `extensions`

Keys:
- `openWithFileSystemByExtension` (object): maps a file extension (lowercase, leading dot like `".zip"`) to a file system plugin ID (example: `"builtin/file-system-7z"`).

Notes:
- The host uses this map when activating a file while browsing the `file` plugin: a matching entry opens the file as a virtual file system instead of `ShellExecute`.
- Default mappings include `.7z`, `.zip`, and `.rar` → `builtin/file-system-7z`.
- To disable auto-mount behavior, set `openWithFileSystemByExtension` to `{}`.
- Viewer extension routing is no longer stored in `extensions`. Schema v16 moved viewer/editor launch configuration into `fileActions`.
- Settings files that still contain root `viewers`, root `editors`, or `extensions.openWithViewerByExtension` are invalid v15-era shapes and are rejected by v16 loading.

## File Actions (v16)

File action settings live under:
- `fileActions.viewers`
- `fileActions.editors`

User Menu action settings live under:
- `userMenu.actions`

The model separates two concepts:
- **Actions** are named things RedSalamander can launch.
- **Associations** are command rules that choose an action for a file match and optional computer name.

Viewer action settings:
- `fileActions.viewers.actions` (array): viewer actions shown by `View With` and referenced by viewer associations.
- `fileActions.viewers.associations` (array): ordered association rules for `F3 View` and `Alt+F3 Alternate View`.

Editor action settings:
- `fileActions.editors.actions` (array): external-program editor actions shown by `Edit With`, `Edit New`, and referenced by editor associations. Editor actions MUST use `kind: "externalProgram"`; `viewerPlugin` actions are valid only in `fileActions.viewers.actions`.
- `fileActions.editors.associations` (array): ordered association rules for `F4 Edit`, `Ctrl+Shift+F4 Alternate Edit`, and `Shift+F4 Edit New`.

User Menu settings:
- `userMenu.actions` (array): ordered external actions shown by `cmd/pane/userMenu` and launched by `cmd/pane/userMenu/<itemId>`.

Each action definition:
- `id` (string): stable action ID used by settings and command payloads. IDs preserve configured casing when saved or displayed, but comparisons and uniqueness are case-insensitive. IDs must start with an ASCII letter or digit, may contain ASCII letters, digits, `_`, `.`, `-`, and `/`, and must be unique within the containing action array case-insensitively.
- `displayName` (string, optional): localized/user-facing label shown in pickers and preferences.
- `enabled` (bool, optional): default `true`; disabled actions remain configured but cannot be launched.
- `kind` (string): required; `viewerPlugin` for an internal viewer plugin, or `externalProgram` for an external process. Unknown kinds are invalid.
- `pluginId` (string, required when `kind` is `viewerPlugin`, forbidden when `kind` is `externalProgram`): viewer plugin ID.
- `executablePath` (string, required when `kind` is `externalProgram`, forbidden when `kind` is `viewerPlugin`): an explicit absolute drive, UNC, or Win32 extended path to the process. Relative paths, drive-relative paths, rooted paths without a drive, device paths, and bare program names are never resolved through `PATH` or a working directory.
- `arguments` (string, optional): process arguments, which may contain launch macros.
- `workingDirectory` (string, optional): process working directory, which may contain launch macros.
- `appliesTo.matches` (array, optional): file matches the action can handle; empty means no file-match filter.
- `appliesTo.computerNames` (string[], optional): unique non-empty computer-name filters; empty means any computer.

Each file match:
- `kind: "default"` means the catch-all `*` match; `value` must be omitted or empty.
- `kind: "extension"` stores an extension with a leading dot, such as `".txt"` or `".EXE"`, using ASCII letters, digits, `_`, `.`, and `-`.
- `kind: "pattern"` stores a non-empty case-insensitive filename glob, such as `"*.test.log"`, with a maximum length of 512 characters.

Each viewer association rule:
- `match` (object): default, extension, or pattern match.
- `computerName` (string, optional): exact computer name. Empty means any computer.
- `viewActionId` (string, required): action for `F3 View`.
- `alternateViewActionId` (string, optional): action for `Alt+F3 Alternate View`; empty means no configured alternate viewer.

Each editor association rule:
- `match` (object): default, extension, or pattern match.
- `computerName` (string, optional): exact computer name. Empty means any computer.
- `editActionId` (string, required): action for `F4 Edit`.
- `alternateEditActionId` (string, optional): action for `Ctrl+Shift+F4 Alternate Edit`; empty means no configured alternate editor.
- `editNewActionId` (string, optional): action for `Shift+F4 Edit New`; empty means create the file without launching an editor when no applicable edit-new action is chosen.

Pattern and extension association rows share the same specificity bucket. If multiple rows in the same priority bucket match, the first row in persisted order wins.

Association action IDs must reference an action in the same `viewers.actions` or `editors.actions` array. References match action IDs case-insensitively. Two actions whose IDs differ only by case are invalid duplicates and must be rejected by the reader before resolution or `View With` / `Edit With` collection. The reader rejects duplicate association keys with the same `match` plus `computerName`, duplicate action IDs, malformed matches, unknown action kinds, editor actions that are not `externalProgram`, `viewerPlugin` actions without `pluginId`, `viewerPlugin` actions with `executablePath`, `externalProgram` actions without `executablePath`, and `externalProgram` actions with `pluginId`. Normal `LoadSettings` backs up malformed settings and starts from defaults; no-recovery loads return a failure. For compatibility, a configured relative or bare external executable is preserved verbatim but forced disabled instead of invalidating unrelated settings. Preferences blocks newly edited literal external paths unless they are explicit and absolute, but it still allows supported macros in `executablePath`; it never guesses or rewrites a path.

Launch macro strings are interpreted by the command layer, not by the settings store. Supported macros are `{Path}`, `{FullPath}`, `{PathAndFilename}`, `{Filename}`, `{SelectedPathsFile}`, `{OppositePanePath}`, and `{ComputerName}`. `workingDirectory` expands macros as raw text. `arguments` expands each macro as a Windows command-line argument by quoting and escaping the macro value; when a macro is already wrapped in literal quotes in the template, the macro content is escaped without adding another quote pair. Launch planning validates the expanded `executablePath` again and fails unless it is an explicit absolute executable path, so direct callers and externally edited JSON cannot bypass the settings or Preferences checks.

Omitting `fileActions` uses the built-in viewer/editor defaults. Explicit empty `fileActions.viewers` or `fileActions.editors` sections mean the user cleared that action family and must round-trip as empty instead of being repopulated from defaults.

Preferences contract:
- `Preferences -> Viewers` uses `Actions` first, then `Associations`. The Associations table columns are Match, Computer, F3 View, Alt+F3 Alternate View, and Status.
- `Preferences -> Editors` uses `Actions` first, then `Associations`. The Associations table columns are Match, Computer, F4 Edit, Ctrl+Shift+F4 Alternate Edit, Shift+F4 Edit New, and Status.
- Viewer-plugin actions choose `pluginId` through a non-editable viewer-plugin combo that displays plugin names while storing the stable ID; Editors actions do not expose plugin-ID entry.
- Both pages show a resolved-action preview for a test file path using the same resolver priority as the command layer.
- `Preferences -> User Menu` lists configured user-menu external commands and persists ordering, command fields, enabled state, match filters, and computer-name filters.
- These pages MUST mark the dialog dirty when editable action settings change, and `Apply` / `OK` MUST persist the updated `fileActions` or `userMenu` section without dropping unrelated settings.
- Action combo boxes list `(none)` plus configured actions from the matching action family. Selecting `(none)` clears the stored action id.

User Menu runtime contract:
- `cmd/pane/userMenu` opens the dynamic User Menu for the focused pane.
- `cmd/pane/userMenu/<itemId>` launches the configured item whose ID matches case-insensitively when it is enabled and applicable.
- User Menu entries are filtered by enabled state, `appliesTo.matches`, and `appliesTo.computerNames`. Missing external executables are shown disabled in the popup and report a localized unavailable message when launched directly.
- External-program execution uses the shared macro engine and selected-paths-file lifecycle used by viewer/editor actions.

Action resolution precedence for `View`, `Alternate View`, `Edit`, `Alternate Edit`, and `Edit New`:
1. Computer + extension or pattern association.
2. Global extension or pattern association.
3. Computer default association.
4. Global default association.

The resolver returns a structured reason such as `ComputerSpecificMatch`, `GlobalMatch`, `ComputerDefault`, `GlobalDefault`, `NoAssociation`, `MissingAction`, `DisabledAction`, `ActionDoesNotApply`, or `UnsupportedActionKind`. If the resolved action is missing, disabled, filtered out by match/computer, unsupported, or cannot resolve required macro inputs, the command layer must report localized pane feedback instead of silently doing nothing.

## Make File List Settings (v16)

Make File List settings live under:
- `makeFileList`

This section stores the last selected `cmd/pane/makeFileList` dialog options. The writer omits the section when all values are default. When the section is present, the writer omits individual default-valued fields and the reader supplies defaults for missing fields.

Fields:
- `sourceMode` (string): `"selection"` (default) or `"currentFolder"`.
- `recursive` (bool): default `false`.
- `format` (string): `"text"` (default), `"csv"`, or `"json"`.
- `outputTarget` (string): `"clipboard"` (default) or `"file"`.
- `textMacro` (string): default `{fullPath}\t{size}\t{modified}`. Supported row macros are `{filename}`, `{name}`, `{fullPath}`, `{path}`, `{size}`, `{modified}`, `{attributes}`, and `{isDirectory}`.
- `outputFile` (string): optional destination path used when `outputTarget` is `"file"`.
- `includeName` (bool): default `true`.
- `includeFullPath` (bool): default `true`.
- `includeSize` (bool): default `true`.
- `includeModified` (bool): default `true`.
- `includeAttributes` (bool): default `false`.
- `includeDirectories` (bool): default `true`.

## Shortcuts (v11)

Shortcut bindings live under:
- `shortcuts`

Structure:
- `functionBar` (array): Function Bar bindings (typically `F1`..`F12` with modifiers)
- `folderView` (array): FolderView bindings (key chords that apply when a FolderView has focus)
- `functionBarCollapsed` (bool, optional): persisted collapsed state for the Function Bar shortcuts group
- `folderViewCollapsed` (bool, optional): persisted collapsed state for the Folder View shortcuts group
- `sortColumnId` (string, optional): stable logical column ID used for the persisted `ShortcutsWindow` grid sort
- `sortDescending` (bool, optional): persisted sort direction for `sortColumnId`
- `gridLayout` (array, optional): visible `ShortcutsWindow` grid column layout entries, keyed by stable column ID, display index, and width in DIPs

Each binding entry:
- `vk` (string): stable key name (examples: `F1`, `Backspace`, `Tab`, `Enter`, `Space`, `PageUp`, `PageDown`, `Home`, `End`, `Left`, `Right`, `Up`, `Down`, `Insert`, `Delete`, `A`, `0`, `VK_1B`)
- `ctrl` (bool, optional): `true` for Ctrl modifier (omit when `false`)
- `alt` (bool, optional): `true` for Alt modifier (omit when `false`)
- `shift` (bool, optional): `true` for Shift modifier (omit when `false`)
- `commandId` (string): command identifier (must start with `cmd/`); `cmd/shortcut/unassigned` is the internal sentinel for an intentionally unassigned shortcut chord

Notes:
- Command IDs are stable; UI display names are localized resource strings. No user-facing command names are hard-coded in C++.
- Missing shortcut chords mean “use the current canonical default if one exists.” To intentionally remove a default binding, persist the same chord with `commandId: "cmd/shortcut/unassigned"`.
- The `cmd/shortcut/unassigned` sentinel is preserved during save/load and import/export, consumes its chord as a no-op at runtime, is hidden from assignable command lists, and is excluded from command reverse lookup.
- If a binding references a command that is not implemented, invoking it shows a localized “not yet implemented” message box and does nothing else (see `Specs/UI/UI_CommandMenuKeyboard.md`).
- The `ShortcutsWindow` selected row and live search text are transient UI state, not persisted settings fields. Reopen restores persisted group collapse, logical sort, and visible column layout, then selects a valid row from the restored grid.

All persisted grid-layout arrays (`shortcuts.gridLayout`, `search.resultsGridLayout`,
`batchRename.previewGridLayout`, and `fileOperations.issuesPaneGridLayout`) use one parser contract. Each valid
entry requires a non-empty string `columnId`; non-object entries and entries with a missing, wrong-type, or
empty ID are skipped. Optional unsigned `displayIndex` defaults to zero, numeric `widthDip` is clamped to
`0..10000`, and unknown entry members are ignored. Input order and duplicate IDs are preserved for the owning
UI model to validate. Batch Rename additionally strips single-line control characters from `columnId`; the
other sections preserve their decoded IDs. An invalid strict section such as Shortcuts recovers independently
without discarding valid grid layouts in other sections.

## Window Placement

### Stored data

For each top-level window we store:
- `state`: `"normal"` or `"maximized"`
- `bounds`: `{ "x", "y", "width", "height" }` (integers)
- `dpi`: DPI at time of save (integer, optional but recommended)

Window IDs are strings (examples):
- `MainWindow`
- `MonitorWindow`
- `FindFilesWindow`

### Save policy

- Save on application shutdown (primary).
- Optionally debounce-save on meaningful move/resize (future enhancement).

### Restore policy

When restoring a window:
1. Load the saved placement.
2. Convert bounds if needed (DPI change handling, see below).
3. Ensure the window rectangle is **completely visible** on at least one monitor work area.
4. Apply placement (normal bounds first; then maximize if `state == "maximized"`).

### Full-visibility requirement (critical)

When restoring, the window must be entirely within a monitor’s **work area** (taskbar excluded). If not, adjust:

**Algorithm (normative):**
1. Enumerate monitors (`EnumDisplayMonitors`) and collect each `MONITORINFOEXW::rcWork`.
2. If the saved rect is fully contained in any `rcWork`, keep it.
3. Else choose a target work area:
   - Prefer the monitor with the **largest intersection area** with the saved rect.
   - If there is no intersection with any monitor, use the **primary** monitor work area.
4. Clamp size and position so the final rect is fully contained in the chosen `rcWork`:
   - If width/height exceed `rcWork`, shrink to fit (never negative/zero).
   - Clamp `x/y` so `left >= work.left`, `top >= work.top`,
     `right <= work.right`, `bottom <= work.bottom`.

### DPI considerations

Saved bounds are in screen coordinates at the time of saving. To reduce “tiny/huge window” effects after DPI changes:
- Store `dpi` on save (e.g., `GetDpiForWindow()`).
- On restore, scale width/height by `currentDpi / savedDpi` **before** running the visibility clamp.

## Cache Settings (v1)

### Directory Enumeration Cache (DirectoryInfoCache)

RedSalamander maintains an **in-process** cache for directory enumeration results (`IFilesInformation`) so multiple views can reuse the same snapshot without re-enumerating the folder.

Settings live under:
- `cache.directoryInfo`

Supported keys:
- `maxBytes` (integer|string, optional): Hard cap for cached `IFilesInformation` entries (LRU-by-bytes eviction).
  - integer: interpreted as **KiB** (e.g., `1234` = 1234 KiB)
  - string: `"<number><unit>"` with unit `KB|MB|GB` (base 1024, case-insensitive), e.g. `"512MB"`, `"7gb"`
- `maxWatchers` (integer, optional): Maximum number of active folder watchers (change notifications) the cache is allowed to hold.
- `mruWatched` (integer, optional): In addition to pinned/on-screen folders, watch up to this many **MRU** cached folders (best-effort).

Defaults (when keys are missing):
- `maxBytes`: computed from physical RAM at runtime (see `Specs/Core/Core_DirectoryInfoCache.md`)
- `maxWatchers`: implementation default (see `Specs/Core/Core_DirectoryInfoCache.md`)
- `mruWatched`: implementation default (see `Specs/Core/Core_DirectoryInfoCache.md`)

## Theme System (customizable)

### Requirements

- The **current theme** must be stored in configuration.
- Users can define **custom themes** in configuration and select them as current.
- Built-in themes continue to exist in code; config can reference them by ID.
- Windows Contrast Themes (system High Contrast) remain system-controlled; when enabled they override app theme rendering (selection is preserved in config).

### Theme selection

`theme.currentThemeId` is a string:
- Built-in theme IDs:
  - `builtin/system`
  - `builtin/light`
  - `builtin/dark`
  - `builtin/rainbow`
  - `builtin/highContrast` (app-level high contrast theme)
- Custom theme IDs:
  - Must start with `user/` (e.g., `user/solarized-dark`)

### Theme definitions

Custom themes are stored in `theme.themes[]`. Each custom theme:
- Requires `formatVersion: 2`; missing versions and every other version are rejected
- Has an `id` and `name`
- Declares a `baseThemeId` (built-in theme used as a base)
- May define a reusable `palette` map
- Provides a `colors` map of authored semantic overrides

There is no version-1 reader, legacy direct-color-only shape, or flattened compatibility writer. All inline, standalone, imported, exported, and shipped themes use version 2.

RedConfigure standalone export preserves that authored representation with stable key ordering and deterministic comments for the named palette and each semantic color-key group. It writes through a sibling temporary file, atomically replaces the destination, and reparses the written version 2 document before reporting success.

Inline settings themes use a lenient recovery policy so one damaged theme cannot make the entire settings document unreadable:
- Structurally valid theme objects are retained when possible. Names are clamped to the supported length, unknown `baseThemeId` values are preserved for forward compatibility, and invalid known color overrides are ignored while valid fields continue to load.
- Structurally unusable or non-object `theme.themes[]` entries are preserved as opaque JSON entries and written back unchanged by the settings store instead of being silently discarded.
- The settings loader emits one aggregate recovery warning for skipped or repaired inline-theme data rather than logging one warning per field or entry.
- Inline theme IDs are case-sensitive. The first exact ID is active; later exact duplicates and structurally unusable entries retain their authored JSON value and original `theme.themes[]` array position as opaque repair data. Canonical save must re-emit that data instead of silently deleting it. IDs that differ only by case remain distinct because selection, inheritance, and editing use the same case-sensitive identity contract.
- Standalone `Themes\\*.theme.json5` files remain strict: malformed files fail that file's load and do not use the inline-settings recovery policy.

### UI integration (v1)

- `RedSalamander` exposes theme selection in `View → Theme`.
- The menu includes built-in themes and any `user/*` themes found in `theme.themes[]`.
- Built-in menu structure/labels must be declared in `.rc` resources; runtime code only appends dynamic theme entries (see `Specs/Core/Core_Localization.md`).

### Theme files (predefined / disk)

In addition to `theme.themes[]` stored in the settings file, `RedSalamander` may load extra theme definitions from disk:
- Location: `Themes\\*.theme.json5` next to the executable (`RedSalamander.exe`)
- Format: a single `ThemeDefinition` JSON5 object (same shape as items in `theme.themes[]`)
- Persistence: these themes are **not** written into the settings file on save; only `theme.currentThemeId` is persisted when the user selects one
- Precedence: if a theme ID exists both in settings and on disk, the settings version wins
- Both `theme.themes[]` in the settings file and disk theme files must use the shared `ThemeDefinition` JSON5 parser and validation rules so duplicate IDs, invalid IDs, malformed color keys, malformed color values, and schema errors cannot diverge between the two load paths.

### Version 2 color sources

`palette` and `colors` values are authored strings. A value is either `#RRGGBB`, `#AARRGGBB`, or one of the closed expression forms below. References use `palette.<name>` for palette entries and the normal semantic key for `colors` entries. Palette names and semantic keys are case-sensitive; palette names cannot contain `.`.

- `ref(key)`
- `lighten(key,amount)`, `darken(key,amount)`, `alpha(key,amount)`
- `blend(firstKey,secondKey,amount)`
- `contrast(backgroundKey)` or `contrast(backgroundKey,lightKey,darkKey)`
- `perceptualTone(key,0..100)`
- `ensureContrast(foregroundKey,backgroundKey,1..21)`
- `harmonize(key,targetKey,amount)`
- `systemAccent()` or `systemColor(accent|accentLight|accentDark|window|windowText|highlight|highlightText)`
- `tone(lightKey,darkKey)`
- `seededRainbow(runtime.seed,saturation,value,alpha,hueOffset)`
- `seededChoice(runtime.seed,key1,key2[,key3...key8])`

Amounts accept `0.0` through `1.0` or percentage syntax. Expressions are deliberately non-nested; authors name intermediate palette entries instead. Missing references, dependency cycles, references to paint-time sources, and sources longer than 256 UTF-16 code units are validation errors. Resolution has a maximum dependency depth of 32. Theme definitions are bounded to 128 palette entries and 512 semantic entries.

Formatter output uses exactly the camelCase spellings listed above and must remain accepted by `SettingsStore.schema.json`; parsing remains case-insensitive for compatibility. `ensureContrast` evaluates WCAG contrast from the rendered foreground composite, not from its uncomposited RGB channels. Its background must be opaque because the expression has no surface-backdrop argument. Foreground alpha is preserved only when some candidate at that alpha can meet the requested ratio; otherwise resolution fails instead of certifying an inaccessible color.

Static and event-time values resolve when a theme is applied or a relevant system event occurs. Paint-time sources are parsed and compiled once, then evaluated from a stable 32-bit runtime seed without parsing, allocation, locking, or I/O in the paint path. `seededChoice` has two through eight pre-resolved candidates. High Contrast continues to override authored and Rainbow output.

Startup, explicit theme selection, system color/theme notifications, and settings hot reload all resolve the authored graph before application. A hot reload whose selected theme graph does not resolve is rejected before replacing runtime settings, retains the last valid live theme, and uses the existing non-destructive invalid-reload alert path.

`builtin/rainbow` remains the application-wide Rainbow base and preserves its existing light/dark base selection. A version 2 custom theme can inherit it and statically override individual semantic tokens. Static overrides suppress Rainbow only for those tokens. Other bases may use `seededRainbow` or `seededChoice` on supported dynamic tokens without enabling plugin-wide Rainbow behavior.

### Theme performance contract

The protected resolver scenario contains 128 palette entries and 512 semantic sources, including the maximum eight-candidate `seededChoice` program. The protected paint scenario evaluates both eight-candidate `seededChoice` and `seededRainbow` over fixed seeds. Resolver and preview work must be bounded; compiled paint evaluation must remain allocation-free and perform no parsing, locking, callbacks, system queries, or I/O.

Instrumentation uses `theme.resolve_us`, `theme.resolve.node_count`, `theme.resolve.edge_count`, `theme.dynamic.evaluate_us`, `theme.dynamic.evaluate_count`, and `redconfigure.theme.preview_resolve_us`. The deterministic selftest records at least 200 resolver samples and 250 dynamic batches so p95 is meaningful. Release guards are 16.667 ms for the worst allowed graph and 5 ms for a batch of 2,000 dynamic evaluations; Debug guards are 100 ms and 50 ms respectively. Changes to the resolver, compiled programs, preview recomputation, or dynamic paint consumption require same-machine archived evidence under `Specs/TestRuns/`.

### Color keys (recommended set)

Color keys are dot-separated identifiers. Unknown keys are ignored.

**App-level**
- `app.accent`
- `window.background`

**Menu (`MenuTheme`)**
- `menu.background`
- `menu.text`
- `menu.disabledText`
- `menu.selectionBg`
- `menu.selectionText`
- `menu.separator`
- `menu.border`

**NavigationView (`NavigationViewTheme`)**
- `navigation.background`
- `navigation.backgroundHover`
- `navigation.backgroundPressed`
- `navigation.text`
- `navigation.separator`
- `navigation.accent`
- `navigation.progressOk`
- `navigation.progressWarn`
- `navigation.progressBackground`

**FolderView (`FolderViewTheme`)**
- `folderView.background`
- `folderView.itemBackgroundNormal`
- `folderView.itemBackgroundHovered`
- `folderView.itemBackgroundSelected`
- `folderView.itemBackgroundSelectedInactive`
- `folderView.itemBackgroundFocused`
- `folderView.textNormal`
- `folderView.textSelected`
- `folderView.textSelectedInactive`
- `folderView.textDisabled`
- `folderView.focusBorder`
- `folderView.gridLines`
- `folderView.errorBackground`
- `folderView.errorText`
- `folderView.warningBackground`
- `folderView.warningText`
- `folderView.infoBackground`
- `folderView.infoText`

**File Operations Popup**
- `fileOps.progressBackground`
- `fileOps.progressTotal`
- `fileOps.progressItem`
- `fileOps.graphBackground`
- `fileOps.graphGrid`
- `fileOps.graphLimit`
- `fileOps.graphLine`
- `fileOps.scrollbarTrack`
- `fileOps.scrollbarThumb`

**ViewerText diff surfaces**
- `viewer.diff.addedBackground`
- `viewer.diff.removedBackground`
- `viewer.diff.contextBackground`
- `viewer.diff.headerBackground`
- `viewer.diff.bannerBackground`
- `viewer.diff.placeholderBackground`
- `viewer.diff.divider`

These `viewer.diff.*` keys are semantic app-theme tokens used by `builtin/viewer-text`
for parsed diff rendering across light, dark, built-in `builtin/rainbow` / `ThemeMode::Rainbow`,
custom JSON5 themes, and high-contrast themes.
They are part of the app theme contract and must not be stored inside
`settings.plugins.configurationByPluginId["builtin/viewer-text"]`.
ViewerText must not introduce separate persisted diff visual-style keys for row coloring,
banner presentation, or active hunk/section state.

This list is the initial contract; additional keys may be added over time.

### Theme color key maintenance (mandatory)

Whenever you add/remove/rename a theme color key (or change its semantics), update all of:
- `Specs/SettingsStore.schema.json` (settings validation + editor IntelliSense)
- `Specs/Core/Core_SettingsStore.md` (this contract list)
- Built-in theme defaults in `RedSalamander/AppTheme.h` / `RedSalamander/AppTheme.cpp`
- Theme override mapping and fallbacks in `RedSalamander/RedSalamander.cpp` (`ApplyThemeOverrides`)
- Shipped theme files (`Specs/Themes/*.theme.json5` and `Themes\\*.theme.json5` next to `RedSalamander.exe` if present)
- Any component spec that references the token (e.g. `Specs/UI/UI_FolderView.md`)

### Monitor (RedSalamanderMonitor / ColorTextView)

Color keys for the monitor text view. These map to `ColorTextView::Theme`.

- `monitor.textView.bg`
- `monitor.textView.fg`
- `monitor.textView.caret`
- `monitor.textView.selection`
- `monitor.textView.searchHighlight`
- `monitor.textView.gutterBg`
- `monitor.textView.gutterFg`
- `monitor.textView.metaText`
- `monitor.textView.metaError`
- `monitor.textView.metaWarning`
- `monitor.textView.metaInfo`
- `monitor.textView.metaPerf`
- `monitor.textView.metaDebug`

## File Operations Settings

Host-owned File Operations settings live under:

- `fileOperations`

These keys are the global defaults edited by `Preferences -> File Operations`. Plugin-specific knobs such as concurrency, recycle-bin batching, and search walkers are not stored here; they remain in each plugin's own configuration payload.

### Stored data

- `fileOperations.autoDismissSuccess`: whether completed successful/cancelled task cards auto-dismiss from the File Operations popup (bool, default: `false`).
- `fileOperations.popupFooterOnly`: whether the File Operations popup presents only its aggregate footer instead of task cards (bool, default: `false`; owned by the popup, not edited in Preferences).
- `fileOperations.popupCompactDensity`: whether the File Operations popup uses compact task-card density (bool, default: `false`; owned by the popup, not edited in Preferences).
- `fileOperations.preCalcEnabled`: whether copy, move, and permanent delete tasks run the recursive pre-calculation pass when that operation type supports it (bool, default: `true`).
- `fileOperations.preCalcMaxWorkers`: maximum worker count for the host-side pre-calculation tree walk (integer, `1..8`, default: `4`).
- `fileOperations.crossFsBridgeBufferSizeKB`: default per-buffer size for host-driven cross-filesystem bridge copies (integer, `512..16384`, default: `4096`). Two buffers are allocated per active bridged file transfer.
- `fileOperations.defaultBandwidthLimitBytesPerSecond`: default speed limit applied to newly created copy/move tasks when the caller did not already specify one (integer, `>= 0`, default: `0`; `0` means unlimited).

### Diagnostics retention

These keys control the host-owned File Operations diagnostics log and exported issue reports. `maxDiagnosticsLogFiles`, `diagnosticsInfoEnabled`, and `diagnosticsDebugEnabled` surface in `Preferences -> Advanced` (section "File Operations"); the remaining five are advanced JSON-only knobs that are not exposed in Preferences.

- `fileOperations.maxDiagnosticsLogFiles`: maximum number of per-day file-operation diagnostics log files kept on disk (integer, `1..365`, default: `14`).
- `fileOperations.diagnosticsInfoEnabled`: whether info-level entries are written to the per-day diagnostics log (bool, default: `false`; Debug builds default to `true`, Release builds default to `false`).
- `fileOperations.diagnosticsDebugEnabled`: whether debug-level entries are written to the per-day diagnostics log (bool, default: `false`; Debug builds default to `true`, Release builds default to `false`).
- `fileOperations.maxIssueReportFiles`: maximum number of exported file-operation issue reports kept on disk (integer, `>= 1`, default: `60`; advanced JSON-only setting).
- `fileOperations.maxDiagnosticsInMemory`: maximum number of diagnostics entries retained in memory (integer, `>= 1`, default: `256`; advanced JSON-only setting).
- `fileOperations.maxDiagnosticsPerFlush`: pending diagnostics batch threshold before a periodic flush (integer, `>= 1`, default: `64`; advanced JSON-only setting).
- `fileOperations.diagnosticsFlushIntervalMs`: periodic diagnostics flush interval in milliseconds (integer, `>= 1`, default: `5000`; advanced JSON-only setting).
- `fileOperations.diagnosticsCleanupIntervalMs`: interval between diagnostics log cleanup passes in milliseconds (integer, `>= 1`, default: `900000`; advanced JSON-only setting).

### Issues pane view state

- `fileOperations.issuesPaneSortColumnId`: stable column identifier used to restore the issues-pane sort column (string, default: empty).
- `fileOperations.issuesPaneSortDescending`: whether the restored issues-pane sort direction is descending (bool, default: `false`). The value remains part of non-default-state detection even when a malformed or migrated payload omits the corresponding column id.
- `fileOperations.issuesPaneGridLayout`: persisted issues-pane column order and widths (array of grid-column layout entries, default: empty).

### UI ownership

- `Preferences -> File Operations` edits only the host-owned operational defaults above. Popup presentation and issues-pane view-state keys are written by their owning UI surfaces.
- `Preferences -> Plugins -> File System` owns plugin-specific file-operation settings such as:
  - `concurrencyMode`
  - `copyMoveMaxConcurrency`
  - `deleteMaxConcurrency`
  - `deleteRecycleBinMaxConcurrency`
  - `recycleBinBatchSize`
  - `searchMaxDirectoryWalkers`

## Folders (multi-pane)

### Stored data

Folder state is stored as an array to support multiple panes (e.g., Left/Right) and future layouts.

- `folders.active`: the active pane slot (string, recommended: `"left"` or `"right"`).
- `folders.layout.splitRatio`: divider position as a fraction of total width (number, `0.0..1.0`, default: `0.5`).
- `folders.layout.zoomedPane`: which pane is maximized (string; `"left"` or `"right"`; omitted when not maximized).
- `folders.layout.zoomRestoreSplitRatio`: the ratio to restore when un-maximizing (number, `0.0..1.0`; omitted when not maximized).
- `folders.showHiddenFiles`: whether hidden files and folders are displayed (bool, default: `true`).
- `folders.showSystemFiles`: whether system files and folders are displayed (bool, default: `true`).
- `folders.historyMax`: maximum number of stored history entries (integer, default: `20`, clamped to `1..50`).
- `folders.history`: global MRU list of recently visited locations (same format as `current`), most recent first.
- `folders.historyFilters`: per-history-entry filter state for `cmd/pane/filter` (array, optional). Each item:
  - `path` (string): the matching history path (same format as `folders.history[]`)
  - `enabled` (bool): whether filtering is enabled for this path
  - `text` (string): filter mask text (same syntax as selection-mask dialogs)
- `folders.items[]`: per-pane state objects:
  - `slot`: pane identifier / position (string, recommended: `"left"` / `"right"`).
  - `current`: current location as a string (either a Windows path or a plugin-qualified path: `<pluginShortId>:<pluginPath>`).
  - `view.display`: `"brief" | "detailed" | "extraDetailed" | "thumbnails"` (default: `"brief"`). Thumbnails is an exclusive display mode, not a flag layered on top of another display mode.
  - `view.thumbnailSizeDip`: thumbnail visual size in device-independent pixels for that pane. Allowed values are `48`, `64`, `96`, and `128`; missing or invalid values load as `64`, and writers omit the value when it is `64`.
  - `view.sortBy`: `"none" | "name" | "extension" | "time" | "size" | "attributes"` (default: `"name"`).
  - `view.sortDirection`: `"ascending" | "descending"` (default: `"ascending"`; Time/Size typically select `"descending"`).
  - `view.fileExtensionsVisible`: whether file extensions are shown in that pane's display labels (bool, default: `true`).
  - `view.thumbnailsVisible`: deprecated compatibility flag. Readers treat `true` as `view.display: "thumbnails"`; writers persist thumbnail mode through `view.display` and omit this flag.
  - `view.navigationBarVisible`: navigation/address bar visibility for that pane (bool, default: `true`).
  - `view.filterBarVisible`: persistent filter bar visibility for that pane (bool, default: `false`).
  - `view.statusBarVisible`: status bar visibility for that pane (bool, default: `true`).

### UI behavior (v1)

- FolderWindow stores the splitter position in `folders.layout.splitRatio` while dragging; the splitter can be moved all the way to either edge (no minimum pane width); double-clicking the splitter resets it to `0.5`.
- Per-pane `view.*` values reflect the pane’s **Sort by** / **Display as** selections and pane view options for file-extension labels, thumbnail size, thumbnails display mode, navigation bars, filter bars, and status bars.

**History rules (normative):**
- Folder history is global (shared across panes) and stored in `folders.history`.
- `folders.historyMax` defaults to `20` and is clamped to `1..50`.
- Deduplicate paths (if a path already exists, move it to the front).
- `folders.history[0]` is the most recently visited location (from either pane).

## Search Dialog State

Find dialog settings live under:
- `search`

Keys map to `Common::Settings::SearchDialogSettings`:
- `recentRoots` (array, max 10): most-recent-first root history
- `recentNamePatterns` (array, max 10): most-recent-first name-pattern history
- `recentContentPatterns` (array, max 10): most-recent-first content-pattern history
- `lastRoot` (string): last root value shown in the dialog
- `lastNamePattern` (string): last name-pattern value
- `lastContentPattern` (string): last content-pattern value
- `recursive` (bool, default: `true`)
- `includeFiles` (bool, default: `true`)
- `includeDirectories` (bool, default: `false`)
- `followSymlinks` (bool, default: `false`)
- `matchCaseName` (bool, default: `false`)
- `matchCaseContent` (bool, default: `false`)
- `preferIndex` (bool, default: `true`)
- `wantSnippets` (bool, default: `false`)
- `nameMode` (`"wildcard" | "literal" | "regex"`, default: `"wildcard"`)
- `contentMode` (`"disabled" | "textLiteral" | "textRegex"`, default: `"disabled"`)
- `maxResults` (integer, default: `0` meaning unlimited)

Behavior:
- History lists are deduplicated and capped at 10 items per field.
- The host persists the last-used options when the Find window closes or starts a new search.
- Window placement for the Find dialog is stored separately under `windows.FindFilesWindow`.

## Batch Rename Dialog State

Batch Rename settings live under:
- `batchRename`

Keys map to `Common::Settings::BatchRenameSettings`:
- `lastRoot` (string): last Batch Rename root path shown by the window
- `recentMasks` (array, max 10): most-recent-first mask history for folder-scope target collection
- `recentNameTemplates` (array, max 10): most-recent-first new-name template history
- `recentSearchPatterns` (array, max 10): most-recent-first search pattern history
- `recentReplacePatterns` (array, max 10): most-recent-first replacement pattern history
- `includeSubdirectories` (bool, default: `false`)
- `includeFiles` (bool, default: `true`)
- `includeFolders` (bool, default: `false`)
- `regexEnabled` (bool, default: `false`)
- `caseSensitive` (bool, default: `true`)
- `wholeWords` (bool, default: `false`)
- `replaceOnce` (bool, default: `false`)
- `excludeExtension` (bool, default: `false`)
- `flattenSeparator` (string, default: `" - "`): separator used by flattened relative-folder macros
- `fileNameCaseStyle` (`"none" | "lower" | "upper" | "mixed"`, default: `"none"`)
- `extensionCaseStyle` (`"none" | "lower" | "upper" | "mixed"`, default: `"none"`)
- `previewSortColumnId` (string, optional): stable preview-grid column ID used for persisted sort
- `previewSortDescending` (bool, default: `false`)
- `previewGridLayout` (array, optional): preview-grid column layout entries, keyed by stable column ID, display index, and width in DIPs

Behavior:
- History lists are deduplicated case-insensitively, stripped of hidden/control characters, and capped at 10 items per field.
- The Batch Rename window restores persisted rule defaults before visible controls are created.
- On close, the Batch Rename window persists the current root path, mask/rule histories, scope flags, option flags, case transforms, preview sort, and preview grid layout.
- Manual-mode multiline names are transient and MUST NOT be persisted or folded into template/search/replace history.
- Window placement for the Batch Rename window is stored separately under `windows.BatchRenameWindow`.

### Restore behavior

On startup:
- If a pane `current` is missing or invalid/unavailable, fall back to a safe default:
  - `FOLDERID_Documents` (preferred) or another standard folder.
- If `folders.active` does not match any `folders.items[].slot`, select the first item.

## RedSalamander Main Menu State

These settings persist the visibility of the main application menu bar.

Settings live under:
- `mainMenu`

Keys:
- `menuBarVisible` (bool, default: `true`): whether the menu bar is visible.
- `functionBarVisible` (bool, default: `true`): whether the bottom Function Bar is visible.

Behavior:
- When `menuBarVisible` is `false`, pressing **Alt** (alone) temporarily shows the menu bar for interaction; it hides again when the menu loop exits.
- `functionBarVisible` is toggled by `View → Function Bar` and takes effect immediately.

## DxUI Customization State

These settings persist the app-wide DxUI customization values edited in `Preferences -> General -> DxUI`.

Settings live under:
- `ui`

Keys:
- `compactMode` (bool, default: `false`): whether supported DxUI surfaces use compact density instead of standard density.
- `language` (string, default: `"system"`): application language preference. `"system"` follows the Windows preferred UI language list; concrete values are BCP-style language tags such as `"fr"` or `"fr-FR"`.
- `reducedMotion` (string, default: `"system"`): `"system" | "on" | "off"`.
- `windowBackdrop` (string, default: `"default"`): `"default" | "none" | "mica" | "micaAlt" | "acrylic"`.

Behavior:
- `compactMode` is the persisted app-wide default for supported DxUI density-aware surfaces such as Preferences pages, menu bars, popup menus / flyouts, combo boxes, Monitor chrome, and shared grid/list surfaces that opt into the shared density contract.
- `language` drives runtime localization resource lookup for registered application, monitor, and plugin resource owners. Missing, empty, or invalid values fall back to `"system"`.
- Saving omits `language` when it is the default `"system"` and writes the concrete culture tag for non-default selections.
- `reducedMotion` overrides the OS reduced-motion preference for DxUI animation behavior when set to `"on"` or `"off"`. `"system"` follows the OS setting.
- `windowBackdrop` stores the requested DWM backdrop policy for supported top-level windows. Unsupported environments, failed DWM application, and high-contrast mode fall back to `None`.
- When a supported top-level window uses a non-`None` backdrop, its title-bar caption/border/text colors fall back to the DWM system defaults instead of forcing an app-specific title-bar color.
- Preferences previews these values live through `workingSettings`, but only `Apply` / `OK` persist them to the settings store.

## Compare Directories Defaults

Defaults for the Compare Directories feature live under:
- `compareDirectories`

Keys map to `Common::Settings::CompareDirectoriesSettings`:
- Compare files: `compareSize`, `compareDateTime`, `compareAttributes`, `compareContent`
- Content compare: `contentCompareWorkerCount` (integer, `0` = auto (≤4), `1..4` = fixed worker count)
- Subdirectories: `compareSubdirectories`, `compareSubdirectoryAttributes`, `selectSubdirsOnlyInOnePane`
- Ignore patterns: `ignoreFiles` + `ignoreFilesPatterns`, `ignoreDirectories` + `ignoreDirectoriesPatterns`
- Display: `keepIdenticalItems`, `showIdenticalItems`

These defaults are edited in **Preferences → Compare Directories** and are used by the Compare Directories window.

## Connections (Connection Manager)

Global Connection Manager settings live under:
- `connections`

Keys map to `Common::Settings::ConnectionsSettings`:
- `bypassWindowsHello` (bool, default: `false`): when `true`, Windows Hello verification is skipped even if a connection profile requires it (intended for automation).
- `windowsHelloReauthTimeoutMinute` (integer, default: `10`): how long recent interactive authentication is reused for a given connection id (Windows Hello verification or user-entered secret). `0` = re-ask every time.
  - For long-running background operations (copy/compare), the host reuses successful interactive authentication for the remainder of the app run to avoid mid-operation Windows Hello prompts.
- `items` (array): saved `ConnectionProfile` entries (non-secret fields only).

Notes:
- Secrets are not written to the Settings Store JSON; they are stored in Windows Credential Manager (WinCred) or cached in memory for the current app run.
- Connection Manager `Close` and `Connect` commands must first commit pending profile edits into `settings.connections`, update/delete any staged WinCred secrets, and call `SettingsHotReload::SaveSettingsAndSchema(...)`. `Connect` may then return `S_OK` only after the selected connection name is resolvable from the updated runtime settings.
- Quick Connect remains memory-only: it may be shown and edited in the Connection Manager, but the Settings Store writer excludes its fixed `00000000-0000-0000-0000-000000000001` profile from disk JSON and stores any Quick Connect secrets only in the current process.
- Some connection profiles are intentionally hostless. Today that includes S3 / S3 Table (`host = auto region`) and Google Drive (`host` omitted entirely).

## Hot Paths

Hot Paths (bookmarked folders) live under:
- `hotPaths`

Keys:
- `openPrefsOnAssign` (bool): after assigning a slot via `Ctrl+Shift+<digit>`, open **Preferences → Hot Paths**.
- `slots` (array, max 10): hot path slots (`index 0 = Ctrl+1`, `index 9 = Ctrl+0`). Each slot is either `null` (empty) or:
  - `path` (string): folder path
  - `label` (string, optional): display name (empty = show `path`)
  - `showInMenu` (bool): include this slot in the NavigationView drive/menu dropdown

## Selection Masks

History for selection-mask dialogs lives under:
- `selectionMasks`

Keys:
- `selectHistory` (array, max 10): most-recent-first history for `cmd/pane/selection/selectDialog`
- `unselectHistory` (array, max 10): most-recent-first history for `cmd/pane/selection/unselectDialog`
- `filterHistory` (array, max 10): most-recent-first history shared by the `cmd/pane/filter` dialog and the inline pane filter bar

## RedSalamanderMonitor UI State

These settings persist the state of checkable menu items and filter state.

### Stored data

- `monitor.menu.toolbarVisible` (bool)
- `monitor.menu.lineNumbersVisible` (bool)
- `monitor.menu.alwaysOnTop` (bool)
- `monitor.menu.showIds` (bool)
- `monitor.menu.autoScroll` (bool)
- `monitor.filter.mask` (integer, 0-63)
- `monitor.filter.preset` (string): `"custom" | "errorsOnly" | "errorsWarnings" | "allTypes"`

### Defaults (v1)

- `toolbarVisible`: `true`
- `lineNumbersVisible`: `true`
- `alwaysOnTop`: `false`
- `showIds`: `true`
- `autoScroll`: `true`
- `mask`: `63` (all 6 visible message types; bits 0-5 = Text, Error, Warning, Info, Perf, Debug)
- `preset`: `"custom"`

The `63` default and the `0..63` range come from `Specs/SettingsStore.schema.json` (`$defs.monitorFilterSettings.mask`, default `63`, `maximum: 63`). The runtime default is emitted by RedSalamanderMonitor as `Debug::InfoParam::Type::All` (`= 0x3F` = 63; see `Common/Helpers.h`), used in `RedSalamanderMonitor/Configuration.h` (`filterMask = 0x3F`), `RedSalamanderMonitor/Configuration.cpp`, and `RedSalamanderMonitor/Document.h`.

## JSON Schema

- Canonical schema file: `Specs/SettingsStore.schema.json`
- Writer output must conform to this schema.

### `x-ui-*` vendor extensions (Preferences metadata)

`Specs/SettingsStore.schema.json` carries optional vendor-extension keys (the `x-ui-*` prefix is reserved by JSON Schema for non-validating annotations) that describe how each setting is surfaced in the Preferences dialog. They are advisory metadata only: they do not affect validation or the stored value shape, and unknown consumers ignore them.

`RedSalamander/SettingsSchemaParser.cpp` walks both the top-level `properties` tree and the `$defs` section and turns annotated nodes into Preferences fields:

- `x-ui-pane` (string): the Preferences page the setting belongs to (examples: `"General"`, `"Advanced"`, `"Plugins"`, `"Themes"`, `"Panes"`, `"Viewers"`, `"Editors"`, `"Compare Directories"`, `"Keyboard"`, `"UserMenu"`). A node is only treated as a Preferences field when it carries `x-ui-pane`; nodes without it are skipped (though their nested `properties` are still walked).
- `x-ui-section` (string, optional): a sub-heading within the pane used to group related fields (examples: `"Display"`, `"DxUI"`, `"Startup"`, `"Cache"`, `"File Operations"`). Defaults to no section header when omitted.
- `x-ui-order` (integer, optional): display order within the pane/section, ascending. Defaults to `0` when omitted.
- `x-ui-control` (string, optional): the control to render. Known values include `"toggle"` (boolean switch) and `"custom"` (the field is rendered by bespoke page code rather than the generic schema-driven builder). For nested `properties` the default is `"edit"` (a text/number edit box); for top-level `$defs` entries the default is `"custom"`. Fields marked `x-ui-control: "custom"` are excluded from the generic per-pane field builder (`GetNonCustomFieldsForPane`) because their page owns the rendering.
- `x-ui-hidden` (bool, optional): keep the field in the JSON shape but do not render it in an editor. This annotation is consumed by the plugin-configuration editors (`RedSalamander/ManagePluginsDialog.cpp`, `RedSalamander/Preferences.Plugin.Configuration.cpp`) for aggregated plugin schemas; it is not interpreted by `SettingsSchemaParser.cpp`.

Examples in the canonical schema: `mainMenuSettings.menuBarVisible` uses `x-ui-pane: "General"`, `x-ui-section: "Display"`, `x-ui-order: 10`, `x-ui-control: "toggle"`; `fileOperationsSettings.maxDiagnosticsLogFiles` uses `x-ui-pane: "Advanced"`, `x-ui-section: "File Operations"`; and section-owning objects such as `themeSettings`, `pluginsSettings`, and `foldersSettings` carry `x-ui-pane` + `x-ui-control: "custom"` so a dedicated Preferences page renders them.

When adding or moving a setting that should appear in Preferences, set the appropriate `x-ui-*` annotations in `Specs/SettingsStore.schema.json` so the parser routes it to the correct page; omit `x-ui-pane` for settings that are intentionally JSON-only (for example the advanced `fileOperations` diagnostics knobs above).

### Plugin configuration schema/model/codec

`Common::PluginConfiguration` is the single dependency-layer model and codec used by both Manage Plugins and
Preferences. It parses the plugin `fields` array into the supported `text`, `value`, `bool`/`boolean`, `option`,
and `selection` field types, including defaults, numeric limits, choices, folder browsing, and `x-ui-*`
metadata. Invalid JSON, invalid roots/fields, duplicate field keys or choices, unsupported types, invalid
constraints/defaults, and wrong-typed configuration values produce explicit validation issues rather than
surface-specific parser behavior. Fields with an explicit `x-ui-order` are sorted first while source order
remains stable within equal order.

Configuration parsing starts from schema defaults and applies the shared tolerant coercion rules. Serialization
overlays known field values on a deep copy of the original configuration object: unknown/future members retain
their JSON values and original order, known members retain their position when updated, duplicate known members
are removed, and numeric constraints are enforced. Canceling either editor does not serialize or save. An option
is rendered as a boolean toggle only when it has exactly two choices and their labels or values identify distinct
`on`/`off`, `true`/`false`, or `1`/`0` states; all other options remain choice controls. These semantics must be
identical in Manage Plugins and Preferences.

## Example Settings Files

### `RedSalamander-<Major>.<Minor>.settings.json` (example)

```json
{
  "schemaVersion": 16,
  "windows": {
    "MainWindow": {
      "state": "maximized",
      "bounds": { "x": 100, "y": 100, "width": 1280, "height": 800 },
      "dpi": 96
    }
  },
  "theme": {
    "currentThemeId": "user/solarized-dark",
    "themes": [
      {
        "id": "user/solarized-dark",
        "name": "Solarized Dark",
        "baseThemeId": "builtin/dark",
        "colors": {
          "window.background": "#002B36",
          "app.accent": "#268BD2",
          "folderView.background": "#002B36",
          "folderView.textNormal": "#93A1A1",
          "navigation.background": "#073642",
          "menu.background": "#002B36",
          "menu.text": "#EEE8D5"
        }
      }
    ]
  },
  "folders": {
    "history": ["C:\\\\Windows\\\\System32", "C:\\\\Windows", "C:\\\\", "C:\\\\Users"],
    "items": [
      {
        "slot": "left",
        "current": "C:\\\\Windows\\\\System32"
      },
      {
        "slot": "right",
        "current": "C:\\\\",
        "view": { "display": "thumbnails", "sortBy": "time" }
      }
    ]
  }
}
```

### `RedSalamanderMonitor-<Major>.<Minor>.settings.json` (example)

```json
{
  "schemaVersion": 16,
  "windows": {
    "MonitorWindow": {
      "state": "normal",
      "bounds": { "x": 200, "y": 200, "width": 1200, "height": 800 },
      "dpi": 96
    }
  },
  "theme": {
    "currentThemeId": "builtin/dark"
  },
  "monitor": {
    "menu": {
      "alwaysOnTop": true
    },
    "filter": {
      "mask": 1,
      "preset": "errorsOnly"
    }
  }
}
```
