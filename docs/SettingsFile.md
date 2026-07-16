# Settings File & Advanced Configuration

RedSalamander stores per-user settings as JSON5 under:

- `%LocalAppData%\\RedSalamander\\Settings\\`

Typical filenames:

- Release: `RedSalamander-<Major>.<Minor>.settings.json`
- Debug: `RedSalamander-debug.settings.json`
- Legacy (older builds): `RedSalamander.settings.json`

In Debug builds, RedSalamander first looks for `RedSalamander-debug.settings.json`. If it does not exist, it falls back to the versioned or legacy settings file.

You can open the main settings file from **Preferences -> Advanced -> Open settings file**.
You can open the monitor settings file from **Preferences -> Monitor -> Open monitor settings file**.

## Schemas

Two schema files are relevant:

- `RedSalamander.settings.schema.json` next to your user settings file
- `SettingsStore.schema.json` next to `RedSalamander.exe`

The user-side schema file is the easiest one to use when validating manual edits.

### Schema version and upgrades

Every settings file carries a `schemaVersion` field at its root, and RedSalamander loads only the exact version it was built for. There is no automatic preference migration: the app never rewrites an older file into the current shape.

```json
{
  "schemaVersion": 16
}
```

If a file's `schemaVersion` is missing, not an integer, or outside the supported range on a cold startup, RedSalamander backs up the existing file by renaming it to `*.bad.<UTC timestamp>`, restores the built-in defaults, and continues. Your previous values are preserved on disk in the `.bad.*` backup, but they are not carried forward into the new defaults file. When upgrading across a version bump, treat it as a reset of persisted settings rather than a migration.

For the maintainer-facing rules around bumping the version, see [Settings Store Internals](dev/SettingsStore-Internals.md).

## Safe manual-edit workflow

1. Back up the current settings file.
2. Edit the JSON5 file.
3. Save the file.
4. RedSalamander detects the change and applies it without restart.

Notes:

- JSON5 is accepted, so comments and trailing commas are allowed.
- If the edited file is valid, the running app updates immediately.
- If a settings-related dialog is open with unsaved edits, that dialog prompts to reload from disk or keep editing.
- If the edited file is invalid or uses an unsupported schema version during live reload, the running app keeps its current settings and shows a warning; the file is not renamed or backed up in that live path.
- On a later cold startup / recovery load, an invalid settings file still falls back to defaults and is backed up to `.bad.*`.
- Connection Manager secrets are not stored in this JSON file; only non-secret profile data is kept there.
- Only the main RedSalamander settings JSON is watched in this iteration; RedSalamanderMonitor settings and `Themes\\*.theme.json5` files are not live-watched.

## Common advanced edits

### Disable archive auto-mount

This stops archive/container extensions from opening as virtual file systems automatically:

```json
{
  "extensions": {
    "openWithFileSystemByExtension": {}
  }
}
```

### Configure viewer and editor actions

Viewer and editor launch behavior lives under `fileActions`. Actions describe what can be launched; associations describe which action is used for a file match and command:

```json
{
  "fileActions": {
    "viewers": {
      "actions": [
        {
          "id": "viewer-text",
          "displayName": "Text Viewer",
          "enabled": true,
          "kind": "viewerPlugin",
          "pluginId": "builtin/viewer-text",
          "appliesTo": {
            "matches": [{ "kind": "default" }],
            "computerNames": []
          }
        }
      ],
      "associations": [
        {
          "match": { "kind": "extension", "value": ".log" },
          "viewActionId": "viewer-text",
          "alternateViewActionId": ""
        },
        {
          "match": { "kind": "default" },
          "viewActionId": "viewer-text",
          "alternateViewActionId": ""
        }
      ]
    }
  }
}
```

Editor actions use the same action shape under `fileActions.editors`; editor associations use `editActionId`, `alternateEditActionId`, and `editNewActionId`.

Common built-in viewer IDs:

- `builtin/viewer-text`
- `builtin/viewer-imgraw`
- `builtin/viewer-space`
- `builtin/viewer-pe`
- `builtin/viewer-sqlite`
- `builtin/viewer-vlc`
- `builtin/viewer-web`
- `builtin/viewer-json`
- `builtin/viewer-markdown`

Common built-in file-system IDs:

- `builtin/file-system`
- `builtin/file-system-7z`
- `builtin/file-system-ftp`
- `builtin/file-system-sftp`
- `builtin/file-system-scp`
- `builtin/file-system-imap`
- `builtin/file-system-gdrive`
- `builtin/file-system-onedrive-personal`
- `builtin/file-system-onedrive-business`
- `builtin/file-system-sharepoint`
- `builtin/file-system-s3`
- `builtin/file-system-s3table`
- `builtin/file-system-dummy`

### Configure pane view options

Per-pane view options live in `folders.items[].view`. The most common fields are:

- `display`: `"brief"` or `"detailed"`
- `sortBy`: `"none"`, `"name"`, `"extension"`, `"time"`, `"size"`, or `"attributes"`
- `sortDirection`: `"ascending"` or `"descending"`
- `fileExtensionsVisible`: show file extensions in the pane's displayed labels
- `thumbnailsVisible`: show thumbnail-sized visuals in the pane, using shell thumbnails when available and icon fallback otherwise
- `navigationBarVisible`: show the pane navigation/address bar
- `filterBarVisible`: show the pane persistent filter bar
- `statusBarVisible`: show the pane status bar

### Configure User Menu commands

The `userMenu` section stores external commands shown by `F9` and **Commands -> User Menu**. Entries use the same macro fields as viewer/editor external actions:

```json
{
  "userMenu": {
    "actions": [
      {
        "id": "open-log-tool",
        "displayName": "Open in Log Tool",
        "enabled": true,
        "kind": "externalProgram",
        "executablePath": "C:\\Tools\\LogTool\\LogTool.exe",
        "arguments": "{PathAndFilename}",
        "workingDirectory": "{Path}",
        "appliesTo": {
          "matches": [{ "kind": "extension", "value": ".log" }],
          "computerNames": []
        }
      }
    ]
  }
}
```

Useful macros include `{Path}`, `{FullPath}`, `{PathAndFilename}`, `{Filename}`, `{SelectedPathsFile}`, `{OppositePanePath}`, and `{ComputerName}`.

External-action `executablePath` values must be explicit absolute drive, UNC, or Win32 extended paths, or use
the supported macros above when the expanded result is still explicit and absolute. Bare program names such as
`notepad.exe` and relative paths are preserved for repair but forced disabled; they are never resolved through
`PATH` or the current working directory. Arguments and working directories may also use the macros above.

After editing `fileActions`, User Menu, plugin, or file-system extension association settings outside the app, use **Commands -> Reread Associations** to reload those sections without restarting. The command preserves the current pane folders, rebuilds dynamic action menus, refreshes both panes, and leaves the previous valid runtime settings in place if the file is invalid.

### Configure Make File List defaults

The `makeFileList` section stores the last options selected in **Commands -> Make File List**:

```json
{
  "makeFileList": {
    "sourceMode": "currentFolder",
    "recursive": true,
    "format": "csv",
    "outputTarget": "file",
    "outputFile": "C:\\Reports\\files.csv",
    "textMacro": "{filename}|{size}|{attributes}",
    "includeName": true,
    "includeFullPath": true,
    "includeSize": true,
    "includeModified": true,
    "includeAttributes": false,
    "includeDirectories": true
  }
}
```

`sourceMode` is `"selection"` or `"currentFolder"`. `format` is `"text"`, `"csv"`, or `"json"`. `outputTarget` is `"clipboard"` or `"file"`. Text output expands `{filename}`, `{name}`, `{fullPath}`, `{path}`, `{size}`, `{modified}`, `{attributes}`, and `{isDirectory}` for each generated row.

### Tune file-operations defaults

Some file-operations defaults are stored in the settings file:

```json
{
  "fileOperations": {
    "autoDismissSuccess": true,
    "maxDiagnosticsLogFiles": 14,
    "maxIssueReportFiles": 60
  }
}
```

The values above are the built-in defaults: `maxDiagnosticsLogFiles` defaults to `14` and `maxIssueReportFiles` defaults to `60`. Substitute your own retention values; you only need to keep the fields you actually change.

### Reset Find dialog history and defaults

The Find Files and Directories window stores its recent values and last-used options under `search`.

Removing that section resets the dialog to defaults. If you prefer to keep the section present explicitly, an empty object has the same effect:

```json
{
  "search": {}
}
```

If you also want to reset the dialog window position, remove `windows.FindFilesWindow`.

## Other settings sections

The settings file has additional top-level sections beyond the ones above. Several are documented in detail on their own pages; edit them there in the app rather than by hand where possible:

- `shortcuts` - keyboard shortcut bindings. See [Keyboard Shortcuts](KeyboardShortcuts.md).
- `connections` - saved Connection Manager profiles (non-secret fields only). See [Connections](Connections.md).
- `compareDirectories` - Compare Directories command options. See [Compare Directories](CompareDirectories.md).
- `hotPaths` - the 10 hot-path bookmark slots (`Ctrl+1` .. `Ctrl+0`). See [Navigation & Path Syntax](NavigationAndPaths.md).
- `theme` - active theme and user theme definitions. See [Themes](Themes.md).
- `monitor` - RedSalamanderMonitor window preferences. See [Monitor](Monitor.md).
- `mainMenu` - menu-bar and function-bar visibility (`menuBarVisible`, `functionBarVisible`). See the **View** menu in [Main Window](MainWindow.md).
- `ui.language` - application language override. See [Localization](dev/Localization.md).

The `theme.themes[]` entries and standalone `Themes\*.theme.json5` files require `formatVersion: 2`. They store reusable authored values under `palette` and map semantic application keys under `colors`; values may be literals, references, supported transforms, system sources, or allowlisted stable runtime sources. Missing-version/version-1 themes are rejected, and exports are never flattened to legacy direct colors. Prefer Preferences for selection/import/export and RedConfigure for dependency-aware authoring. See [Themes](Themes.md) for examples and the function overview.

A few sections have no dedicated page. Brief guidance for editing them by hand:

### startup

Startup behavior preferences. The only field is `showSplash`, which controls whether a splash screen appears when startup takes longer than about 300ms:

```json
{
  "startup": {
    "showSplash": false
  }
}
```

### cache

Limits for the directory-info cache. Fields live under `cache.directoryInfo`. `maxBytes` accepts an integer count of KiB or a string such as `"256 MB"`; `maxWatchers` and `mruWatched` bound the folder-watcher and recently-watched lists:

```json
{
  "cache": {
    "directoryInfo": {
      "maxBytes": "256 MB",
      "maxWatchers": 64,
      "mruWatched": 256
    }
  }
}
```

### selectionMasks

Most-recent-first mask histories for the Select, Unselect, and Filter dialogs (up to 10 entries each). Clearing a list resets that dialog's history:

```json
{
  "selectionMasks": {
    "selectHistory": ["*.cpp", "*.h"],
    "unselectHistory": [],
    "filterHistory": []
  }
}
```

### batchRename

Persisted Batch Rename options and pattern histories. The full feature is documented on [Batch Rename](BatchRename.md); the section stores last-used options (such as `includeSubdirectories` and `caseSensitive`) plus recent masks, name templates, and search/replace patterns. Removing the section, or leaving an empty object, resets it to defaults:

```json
{
  "batchRename": {}
}
```

## When manual editing is useful

- Disabling archive auto-mount globally
- Bulk-changing file-action associations
- Inspecting or resetting advanced values not exposed in the UI

For basic reset instructions, see: [Troubleshooting / Reset](Troubleshooting.md)

