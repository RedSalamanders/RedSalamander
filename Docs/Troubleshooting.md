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

For safe manual edits and concrete JSON examples, see: [Settings File & Advanced Configuration](SettingsFile.md)

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

## S3 / S3 Table limitations

- `s3:` supports browsing, reading, uploading/overwriting objects, and deleting objects.
- `s3:` does not support server-side rename/move, and deleting a virtual "folder" / prefix is not supported.
- `s3table:` is browse/read only. It exposes tables as generated `*.table.json` documents.

See: [S3 / S3 Table](S3AndS3Table.md)

## Debug build breaks in `d2d1debug3.dll` on exit

In Debug builds with the Direct2D debug layer enabled, Windows can break into the debugger on shutdown with messages like “D2D DEBUG ERROR - Memory leaks detected.”

RedSalamander closes auxiliary top-level windows automatically when the main window is closed, but if you still see this break it usually indicates a real Direct2D lifetime leak (e.g., an `ID2D1DeviceContext`, bitmap, or brush still referenced at process teardown).

What to do:

- Check the Visual Studio **Output** window for the leaked interface list and reference counts.
- Identify which window owns the leaked resources and ensure its `WM_CLOSE` / `WM_DESTROY` path releases device resources (swap chain target detached, COM pointers reset, etc.).

## Comparing file-operations self-test runs

Selftests write their raw artifacts under `%LOCALAPPDATA%\\RedSalamander\\SelfTest\\last_run\\`.

In Debug builds from a repo checkout, the selftest harness automatically archives the meaningful artifacts under `Specs\\TestRuns\\<ComputerHashName>\\...` after each run. If repo auto-detection fails, the trace includes `ArchiveToRepo: repo root not found; skipping.` — re-run from a repo checkout, set `REDSALAMANDER_REPO_ROOT` to the repo root, or manually copy from `%LOCALAPPDATA%\\RedSalamander\\SelfTest\\last_run\\` into `Specs\\TestRuns\\...`.

Then use `Tools\CompareTestRuns.ps1` to diff two archived runs and spot regressions:

```powershell
# Compare two runs (summary + file changes + case changes)
.\Tools\CompareTestRuns.ps1 `
    Specs\TestRuns\<ComputerHashName>\FileOps\2026-02-27_085402 `
    Specs\TestRuns\<ComputerHashName>\FileOps\2026-02-28_000800

# Include a line-level trace-log diff (capped at 50 lines)
.\Tools\CompareTestRuns.ps1 `
    Specs\TestRuns\<ComputerHashName>\FileOps\2026-02-27_121904 `
    Specs\TestRuns\<ComputerHashName>\FileOps\2026-02-27_141202 `
    -ShowTraceDiff -MaxTraceDiffLines 50
```

The script reports:

| Section | What it shows |
|---------|---------------|
| **Summary** | Passed / failed / skipped counts and duration delta |
| **Files** | Files added, removed, or changed (SHA-256) between the two folders |
| **Cases** | Individual test cases whose status, reason, or duration changed |
| **Trace** | (with `-ShowTraceDiff`) Line-level diff of `trace.txt` / `fileops_trace.txt` |

Run `Get-Help .\Tools\CompareTestRuns.ps1 -Detailed` for parameter documentation.

To see timing evolution over many runs, use `Tools\\AnalyzeTestRuns.ps1`:

```powershell
.\Tools\AnalyzeTestRuns.ps1 Specs\TestRuns\<ComputerHashName>\FileOps -TopN 10
```
