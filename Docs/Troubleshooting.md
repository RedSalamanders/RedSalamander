# Troubleshooting / Reset

## Settings location

RedSalamander stores per-user settings under:

- `%LocalAppData%\\RedSalamander\\Settings\\`

Typical filenames:

- Release: `RedSalamander-<Major>.<Minor>.settings.json`
- Debug: `RedSalamander-debug.settings.json`
- Legacy (older builds): `RedSalamander.settings.json`

Note: Debug builds may fall back to the versioned/legacy settings file if the debug settings file does not exist.

A schema file is written next to it for reference:

- `RedSalamander.settings.schema.json`

## Reset to defaults

1. Close RedSalamander.
2. Rename or delete all `RedSalamander*.settings.json` files in `%LocalAppData%\\RedSalamander\\Settings\\`.
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

## Debug build breaks in `d2d1debug3.dll` on exit

In Debug builds with the Direct2D debug layer enabled, Windows can break into the debugger on shutdown with messages like “D2D DEBUG ERROR - Memory leaks detected.”

RedSalamander closes auxiliary top-level windows automatically when the main window is closed, but if you still see this break it usually indicates a real Direct2D lifetime leak (e.g., an `ID2D1DeviceContext`, bitmap, or brush still referenced at process teardown).

What to do:

- Check the Visual Studio **Output** window for the leaked interface list and reference counts.
- Identify which window owns the leaked resources and ensure its `WM_CLOSE` / `WM_DESTROY` path releases device resources (swap chain target detached, COM pointers reset, etc.).
