# Search Perf Closeout Focused Validation

Date: 2026-04-26
Configuration: Debug x64

## Commands

- `.\build.ps1 -Configuration Debug -ProjectName RedSalamander`
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=search_local_plugin_invalid_regex_reports_single_completion`
- `.\.build\x64\Debug\RedSalamander.exe --fileops-selftest --selftest-case=Phase7_WatcherChurn`

## Results

- Commands: passed 1, failed 0.
- FileOperations: passed 3, failed 0.

Artifacts:

- `commands/results.json`
- `commands/trace.txt`
- `fileops/results.json`
- `fileops/trace.txt`
- `build/msbuild-20260426_144734_122.log`
