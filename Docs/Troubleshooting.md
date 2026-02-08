# Troubleshooting / Reset

## Settings location

RedSalamander stores per-user settings under:

- `%LocalAppData%\\RedSalamander\\Settings\\`

Typical filenames:

- Release: `RedSalamander-<Major>.<Minor>.settings.json`
- Debug: `RedSalamander-debug.settings.json`

A schema file is written next to it for reference:

- `RedSalamander.settings.schema.json`

## Reset to defaults

1. Close RedSalamander.
2. Rename or delete the settings file in `%LocalAppData%\\RedSalamander\\Settings\\`.
3. Start RedSalamander again.

## ViewerWeb (HTML/PDF/Markdown/JSON) does not open

- Ensure **WebView2 Runtime** is installed.
- Ensure `Plugins\\ViewerWeb.dll` is present and not disabled.
- If ViewerWeb is missing/disabled, those extensions fall back to the Text viewer.

## VLC viewer says VLC is required

- Install VLC media player, or
- Set the VLC installation folder in Preferences → Plugins → VLC Viewer.

## Remote file systems keep asking for passwords

- Prefer [Connection Manager](Connections.md) instead of storing defaults in plugin settings.
- If “Save password” is unchecked, the secret is kept session-only and you may be prompted again after restart.

## S3 is read-only

The S3 and S3 Table file systems currently implement browsing and reading, but not uploads/deletes/renames.

