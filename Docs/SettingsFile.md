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

1. Close RedSalamander.
2. Back up the current settings file.
3. Edit the JSON5 file.
4. Start RedSalamander and confirm the change.

Notes:

- JSON5 is accepted, so comments and trailing commas are allowed.
- If the settings file is invalid or uses an unsupported schema version, RedSalamander falls back to defaults and backs up the bad file.
- Connection Manager secrets are not stored in this JSON file; only non-secret profile data is kept there.

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

## When manual editing is useful

- Disabling archive auto-mount globally
- Bulk-changing viewer associations
- Inspecting or resetting advanced values not exposed in the UI

For basic reset instructions, see: [Troubleshooting / Reset](Troubleshooting.md)
