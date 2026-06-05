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

On every successful save, a schema file is written next to the settings file:
- `<AppId>.settings.schema.json`

The settings JSON includes a `$schema` property referencing it (relative path):
- `$schema: "./<AppId>.settings.schema.json"`

Notes:
- `Common.dll` writes the base schema (identical to `Specs/SettingsStore.schema.json`) as a best-effort convenience for manual editing.
- `Common.dll` loads the base schema text from `SettingsStore.schema.json` shipped next to the exe (copied from `Specs/SettingsStore.schema.json` during the build) and caches it in memory.
- `RedSalamander.exe` overwrites that file with an aggregated schema that includes plugin configuration schemas under `plugins.configurationByPluginId[pluginId]` (best-effort).

### UI entry point

`Preferences -> Advanced` exposes a command link that opens `GetSettingsPath(L"RedSalamander")` with the shell default editor for the current user. `Preferences -> Monitor` exposes a separate command link that opens `GetSettingsPath(L"RedSalamanderMonitor")`. Either command may create the target JSON/settings schema first when the file is missing, but invoking the link MUST NOT mark Preferences dirty. The generated `RedSalamander.settings.schema.json` MUST NOT expose the root `monitor` property; the generated `RedSalamanderMonitor.settings.schema.json` MUST expose it.

## Read/Write Requirements

### Encoding

- Settings file (`GetSettingsPath(appId)`): **UTF-8** (with BOM)
- Schema file (`<AppId>.settings.schema.json`): **UTF-8** (no BOM)
- In-memory strings: UTF-16 (`std::wstring` / `std::filesystem::path`) as per project convention

### Atomic saves

Saving must be atomic to prevent partial/corrupt writes:
1. Write JSON to a temp file in the same directory:
   - `<SettingsFileName>.tmp` (e.g., `<AppId>-7.0.settings.json.tmp`)
2. Flush buffers.
3. Replace the target file using an atomic rename/replace operation (Windows `MoveFileExW` with replace/write-through semantics).

### Recovery behavior

For the normal startup / explicit recovery path (`LoadSettings(...)`):
- If loading fails (missing file, unreadable file, invalid JSON, or invalid types), start with defaults.
- If a file existed but was invalid, rename it to a backup for diagnostics:
  - `<SettingsFileName>.bad.<UTC timestamp>`
- Startup callers that need to explain recovery to the user call `LoadSettingsWithRecoveryInfo(...)`. When a previous settings file is backed up and defaults are restored, the app MUST show a localized warning that includes the original settings path, the backup path, and explains that the current run is using default settings. The warning should tell the user to close the app, compare the backup with the new settings file at the original path, and copy back only the settings they still need.

For the non-destructive hot-reload path (`TryLoadSettingsNoRecovery(...)`):
- Missing file returns `S_FALSE`.
- Invalid JSON / invalid types / unsupported schema version returns a failure `HRESULT`.
- The file is left in place; it is **not** renamed to `.bad.*`.
- The caller decides whether to keep current runtime settings, warn the user, or recover in some other way.

### Tolerant reads, canonical writes

- Unknown top-level keys must be ignored (forward compatibility).
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
    HRESULT SaveSettings(std::wstring_view appId, const Settings& settings) noexcept;

    // Writes `<AppId>.settings.schema.json` next to the settings file.
    HRESULT SaveSettingsSchema(std::wstring_view appId, std::string_view schemaJsonUtf8) noexcept;
}
```

`SettingsLoadRecoveryInfo` reports whether the startup recovery path used defaults, whether an existing file was moved to a backup, the original settings path, the backup path, the recovery reason, and the unsupported schema version when that was the failure. `LoadSettings(...)` is the compatibility wrapper for callers that do not need this detail.

### File-stamp helper

`SettingsFileStamp` is a stable identity/value snapshot of the current settings file used by live-reload callers to de-duplicate notifications and suppress self-saves.

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
  - `S_OK`: settings loaded and validated
  - `S_FALSE`: file missing
  - failure `HRESULT`: invalid/unreadable/unsupported file without fallback or backup

## Live Reload Semantics (RedSalamander.exe)

`RedSalamander.exe` live-watches only the main settings file returned by `GetSettingsPath(L"RedSalamander")`.

Watcher rules:
- Detection is event-driven (directory change notification), but only the main settings file stamp is authoritative.
- `RedSalamander.exe` writes its own settings through a shared save helper that refreshes the applied stamp immediately after save.
- A changed stamp already recorded as `lastAppliedStamp` or `lastRejectedStamp` is ignored on the next reload check.
- `Themes\\*.theme.json5` files are **not** watched in this iteration.
- `RedSalamanderMonitor.exe` settings do not participate in this main-app hot-reload flow.

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
- Startup recovery for an unsupported version uses the same `.bad.<UTC timestamp>` backup path and user-facing warning as other destructive recoveries. v15 files are intentionally not migrated to v16.
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
- `executablePath` (string, required when `kind` is `externalProgram`, forbidden when `kind` is `viewerPlugin`): process path, which may contain launch macros.
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

Association action IDs must reference an action in the same `viewers.actions` or `editors.actions` array. References match action IDs case-insensitively. Two actions whose IDs differ only by case are invalid duplicates and must be rejected by the reader before resolution or `View With` / `Edit With` collection. The reader rejects duplicate association keys with the same `match` plus `computerName`, duplicate action IDs, malformed matches, unknown action kinds, editor actions that are not `externalProgram`, `viewerPlugin` actions without `pluginId`, `viewerPlugin` actions with `executablePath`, `externalProgram` actions without `executablePath`, and `externalProgram` actions with `pluginId`. Normal `LoadSettings` backs up malformed settings and starts from defaults; no-recovery loads return a failure.

Launch macro strings are interpreted by the command layer, not by the settings store. Supported macros are `{Path}`, `{FullPath}`, `{PathAndFilename}`, `{Filename}`, `{SelectedPathsFile}`, `{OppositePanePath}`, and `{ComputerName}`. `executablePath` and `workingDirectory` expand macros as raw text. `arguments` expands each macro as a Windows command-line argument by quoting and escaping the macro value; when a macro is already wrapped in literal quotes in the template, the macro content is escaped without adding another quote pair.

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
- Has an `id` and `name`
- Declares a `baseThemeId` (built-in theme used as a base)
- Provides a `colors` map of overrides

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
- The disk-theme loader must use the same `ThemeDefinition` JSON5 parser and validation rules as settings-file themes so duplicate IDs, invalid IDs, malformed colors, and schema errors cannot diverge between the two load paths.

### Color representation

Color values are hex strings:
- `#RRGGBB` (opaque)
- `#AARRGGBB` (alpha + RGB)

RedConfigure may offer authoring-time color expressions such as `ref(app.accent)`, `darken(app.accent,20%)`, or `blend(menu.background,app.accent,16%)`. Until the runtime schema explicitly supports a durable `colorExpressions` object, exported `.theme.json5` files must flatten those expressions into direct `colors` values so existing RedSalamander builds can consume them.

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
- `fileOperations.preCalcEnabled`: whether copy, move, and permanent delete tasks run the recursive pre-calculation pass when that operation type supports it (bool, default: `true`).
- `fileOperations.preCalcMaxWorkers`: maximum worker count for the host-side pre-calculation tree walk (integer, `1..8`, default: `4`).
- `fileOperations.crossFsBridgeBufferSizeKB`: default per-buffer size for host-driven cross-filesystem bridge copies (integer, `512..16384`, default: `4096`). Two buffers are allocated per active bridged file transfer.
- `fileOperations.defaultBandwidthLimitBytesPerSecond`: default speed limit applied to newly created copy/move tasks when the caller did not already specify one (integer, `>= 0`, default: `0`; `0` means unlimited).

### UI ownership

- `Preferences -> File Operations` edits only the host-owned `fileOperations.*` defaults above.
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
- `mask`: `31` (all 5 types)
- `preset`: `"custom"`

## JSON Schema

- Canonical schema file: `Specs/SettingsStore.schema.json`
- Writer output must conform to this schema.

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
