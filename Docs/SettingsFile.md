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

### Change viewer mappings

Map file extensions to viewer plugin IDs. Extensions should be lowercase and include the leading dot:

```json
{
  "extensions": {
    "openWithViewerByExtension": {
      ".log": "builtin/viewer-text",
      ".md": "builtin/viewer-text",
      ".json": "builtin/viewer-json"
    }
  }
}
```

Common built-in viewer IDs:

- `builtin/viewer-text`
- `builtin/viewer-imgraw`
- `builtin/viewer-space`
- `builtin/viewer-pe`
- `builtin/viewer-vlc`
- `builtin/viewer-web`
- `builtin/viewer-json`
- `builtin/viewer-markdown`

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
- Bulk-changing viewer associations
- Inspecting or resetting advanced values not exposed in the UI

For basic reset instructions, see: [Troubleshooting / Reset](Troubleshooting.md)
