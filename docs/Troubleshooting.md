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

## Startup command-line switches

`RedSalamander.exe` accepts a few options at launch:

| Switch | Effect |
|--------|--------|
| `-h`, `--help`, `/?` | Show the help text and exit. |
| `--etw` | Enable Info/Perf/Debug ETW diagnostics in Release builds (Debug and ASan Debug builds enable them by default). |
| `--perf` | Write perf metrics to the default JSONL path (Debug and ASan Debug builds enable this by default). |
| `--perf=PATH` | Write perf metrics to a custom JSONL path. |
| `--crash-test` | Trigger the crash-handler test (deliberately crashes to exercise the crash pipeline). |

The default perf JSONL path is `%LocalAppData%\\RedSalamander\\Perf\\RedSalamander_<timestamp>.jsonl`. For the full diagnostics model, the ETW event schema, and the `--perf` sink, see [Diagnostics: ETW, Debug Logging & Perf](dev/Diagnostics.md).

## Crash quarantine (auto-disable a suspect plugin)

If RedSalamander crashes, it leaves a crash marker (and a minidump/report) under `%LocalAppData%\\RedSalamander\\Crashes\\`. On the next launch, if a marker is present, RedSalamander checks which file-system plugin(s) were active at the time of the crash and offers to disable them:

- A **Crash recovery** prompt asks whether to disable the suspect plugin(s) to avoid repeated crashes.
- Choosing **Yes** adds the plugin id(s) to the disabled list in your settings (and clears the current file-system plugin if it was one of them). Choosing **No** leaves your settings unchanged.

The marker is cleared on a clean shutdown, so the prompt only appears after an actual crash. To re-enable a plugin you disabled this way, remove it from the disabled plugins list in Preferences → Plugins (or edit the settings file — see [Settings File & Advanced Configuration](SettingsFile.md)).

## Capturing diagnostics for a bug report

RedSalamander writes no log files; diagnostics are emitted as ETW (Event Tracing for Windows) events and consumed in real time by `RedSalamanderMonitor.exe`. To capture errors and warnings while reproducing a problem:

1. Start `RedSalamanderMonitor.exe` and let it begin its ETW session.
2. Reproduce the issue in RedSalamander.
3. Copy the relevant `[Error]` / `[Warning]` lines from the Monitor to attach to your report.

In Release builds only **Error** and **Warning** events are emitted by default; Info/Perf/Debug events require launching with `--etw` (see the switch table above). If the Monitor cannot start a session (`ERROR_ACCESS_DENIED`), run `.\\init-etw-trace.ps1` once and sign out/in. See [Monitor](Monitor.md) for the full walkthrough and [Diagnostics: ETW, Debug Logging & Perf](dev/Diagnostics.md) for details.

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
