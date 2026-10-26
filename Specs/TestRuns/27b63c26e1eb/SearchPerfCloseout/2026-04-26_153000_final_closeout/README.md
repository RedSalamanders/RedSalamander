# FileSystem Stability/Search/Perf Final Closeout

Date: 2026-04-26
Configuration: Debug x64

## Commands

- `.\build.ps1 -Configuration Debug -ProjectName RedSalamander`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=search_local_plugin_invalid_regex_reports_single_completion`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=search_local_plugin_parallel_cancel_fanin`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=filesystem_local_watch_unwatch_drains_inflight_callback`
- `.\.build\x64\Debug\RedSalamander.exe --fileops-selftest --selftest-case=Phase7_WatcherChurn`
- `.\.build\x64\Debug\RedSalamander.exe --fileops-selftest --selftest-case=Phase7_CopyMoveConcurrency16Perf`
- `.\.build\x64\Debug\RedSalamander.exe --fileops-selftest --selftest-case=Phase7_SharedPerItemScheduler`
- `.\.build\x64\Debug\RedSalamander.exe --fileops-selftest --selftest-case=Phase11_BridgePipelineDummyToDummyPerf`
- `.\.build\x64\Debug\RedSalamander.exe --fileops-selftest --selftest-case=Phase14_PopupHostLifetimeGuard`

## Results

- Build: passed.
- Commands focused cases: 3 passed, 0 failed.
- FileOperations focused cases: each run passed setup, the focused phase, and cleanup; 0 failed.

Artifacts are stored per case under `commands/` and `fileops/`, with the build log under `build/`.
