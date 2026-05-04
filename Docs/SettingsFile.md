# Settings File & Advanced Configuration

RedSalamander stores per-user settings as JSON5 under:

- `%LocalAppData%\\RedSalamander\\Settings\\`

Typical filenames:

- Release: `RedSalamander-<Major>.<Minor>.settings.json`
- Debug: `RedSalamander-debug.settings.json`
- Legacy (older builds): `RedSalamander.settings.json`

In Debug builds, RedSalamander first looks for `RedSalamander-debug.settings.json`. If it does not exist, it falls back to the versioned or legacy settings file.

## Schemas

Two schema files are relevant:

- `RedSalamander.settings.schema.json` next to your user settings file
- `SettingsStore.schema.json` next to `RedSalamander.exe`

The user-side schema file is the easiest one to use when validating manual edits.

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
- Only the main settings JSON is watched in this iteration; `Themes\\*.theme.json5` files are not live-watched.

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
    "maxDiagnosticsLogFiles": 20,
    "maxIssueReportFiles": 20
  }
}
```

### Reset Find dialog history and defaults

The Find Files and Directories window stores its recent values and last-used options under `search`.

Removing that section resets the dialog to defaults. If you prefer to keep the section present explicitly, an empty object has the same effect:

```json
{
  "search": {}
}
```

If you also want to reset the dialog window position, remove `windows.FindFilesWindow`.

## When manual editing is useful

- Disabling archive auto-mount globally
- Bulk-changing file-action associations
- Inspecting or resetting advanced values not exposed in the UI

For basic reset instructions, see: [Troubleshooting / Reset](Troubleshooting.md)
