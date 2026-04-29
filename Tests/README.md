# RedSalamander Test Infrastructure

## Overview

RedSalamander has **1,162 test cases** across 8 test suites covering UI automation, plugin
integration, file operations, DirectX rendering, performance baselines, and ETW diagnostics.

| Category | Suites | Tests | Framework |
|----------|--------|-------|-----------|
| Self-Tests (in-process) | 3 | 567 | Custom `SelfTest::RunCase` |
| DxUi Component Tests | 1 | 571 | Standalone harness |
| Performance Tests | 1 | 4 | CppUnitTest DLL |
| Viewer Plugin Tests | 2 | 17 | Standalone harness |
| Monitor/ETW Tests | 1 | 3 | Standalone harness |

Related specifications:
- `Specs/Testing/Testing_SelfTests.md` — result contract
- `Specs/Testing/Testing_TestCoverage.md` — per-case inventory
- `Specs/Testing/Testing_PerformanceValidation.md` — perf validation

## Quick Start

```powershell
.\Tools\Run-AllTests.ps1                      # Build Debug + run ALL self-tests
.\Tools\Run-AllTests.ps1 -Suite Commands       # Single suite
.\Tools\Run-AllTests.ps1 -SkipBuild -FailFast  # Fast iteration
.\Tools\Run-AllTests.ps1 -TimeoutMultiplier 3  # Slow machines
```

---

## 1. Commands Self-Test Suite — 346 cases

**Flag:** `--commands-selftest`
**Source:** `RedSalamander\SelfTest\Commands\Commands.SelfTest.cpp` + 11 `.cpp` family files

Tests UI automation, dialog interactions, preferences, shortcuts, themes, navigation,
and command dispatch inside the live application window.

| Family | File | Tests | Coverage |
|--------|------|-------|----------|
| Settings | `SelfTest\Commands\Commands.SelfTest.Settings.cpp` | 11 | Hot-reload, registry store, shortcut defaults |
| PluginConfig | `SelfTest\Commands\Commands.SelfTest.PluginConfig.cpp` | 10 | Plugin configuration, file-system plugin |
| Connections | `SelfTest\Commands\Commands.SelfTest.Connections.cpp` | 17 | Connection manager, credential prompts |
| Preferences | `SelfTest\Commands\Commands.SelfTest.Preferences.cpp` | 129 | Preferences dialog — all categories, DxUi surfaces |
| CompareOptions | `SelfTest\Commands\Commands.SelfTest.CompareOptions.cpp` | 8 | Compare directories options, progress |
| Search | `SelfTest\Commands\Commands.SelfTest.Search.cpp` | 52 | Find dialog, local search index, quick search/filter |
| Shortcuts | `SelfTest\Commands\Commands.SelfTest.Shortcuts.cpp` | 24 | Shortcuts window, key binding |
| ViewCommands | `SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp` | 27 | View commands, selection, sort, pane, tabs |
| FileOps | `SelfTest\Commands\Commands.SelfTest.FileOps.cpp` | 19 | File operations issues pane, speed limit |
| Navigation | `SelfTest\Commands\Commands.SelfTest.Navigation.cpp` | 2 | Navigation location, GoTo |
| Dialogs | `SelfTest\Commands\Commands.SelfTest.Dialogs.cpp` | 47 | About, fatal error, splash, change case, rename, filter, mask |

## 2. Compare Directories Self-Test Suite — 145 cases

**Flag:** `--compare-selftest`
**Source:** `RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.cpp` + 3 included case files

Tests the compare-directories engine including local/remote search, indexing, and session logic.

| Family | Tests | Coverage |
|--------|-------|----------|
| Core session | ~25 | Unique, size, content, unicode, subdirs, ignore patterns, invalidation |
| Search service | 25 | CLI, bootstrap, cold start, compact, maintenance, multi-client |
| Local search | 18 | QI, callbacks, backend prefs, tree walking, wildcards, content |
| Local index core | 12 | Snapshot reload, journal replay, corruption rebuild, SQLite |
| SQLite index store | 6 | Schema bootstrap, compaction, checkpoint, upgrade |
| Host fallback | 7 | Plugin path root, degraded IO, cancel, access denied |
| Remote filesystems | 9 | S3, OneDrive, SharePoint, FTP directories and size callbacks |
| Google Drive | 3 | Plugin contract, client ID, refresh token |
| OAuth / credentials | 3 | Token storage, auth mode, Windows Hello cache |
| Directory size | 3 | Local, dummy, 7z filesystem callbacks |
| Search text helpers | 2 | Text matching, decoding |
| Misc (concurrency, caching, UI) | ~32 | Crash quarantine, setCompareEnabled, uiVersion, etc. |

## 3. File Operations Self-Test Suite — 76 phases

**Flag:** `--fileops-selftest`
**Source:** `RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.SelfTest.cpp` + 4 included phase files

Async tick-driven state machine testing file copy/move/delete operations end-to-end.
Organised into 12 families spanning phases 5–16.

| Family | Phases | Coverage |
|--------|--------|----------|
| Phase 05 — PreCalc | 7 | Pre-calculation settings, cancel, latency, mode switching |
| Phase 06 — PopupAndDelete | 4 | Popup, bandwidth throttle, delete operations |
| Phase 07 — WatchAndParallelism | 14 | Watchers, cache, parallelism, concurrency |
| Phase 08 — Validation | 5 | Defaults, destinations, size bytes |
| Phase 09 — ConflictPrompt | 7 | Overwrite, apply-to-all, skip, retry |
| Phase 10 — DeleteValidation | 1 | Permanent delete |
| Phase 11 — BridgeAndConnections | 7 | Cross-filesystem bridges, connection overrides |
| Phase 12 — Reparse | 1 | Reparse point policy |
| Phase 13 — PostMortem | 1 | Post-mortem diagnostics |
| Phase 14 — PopupLifetime | 1 | Popup lifetime guard |
| Phase 15 — FileSystem7z | 2 | 7z filesystem operations |
| Phase 16 — Remote | 19 | FTP, SFTP, SCP, IMAP, S3, OneDrive, SharePoint |

---

## 4. DxUi Component Tests — 571 tests

**Project:** `Tests\DxUiTests\`  •  **Run:** `.\.build\x64\Debug\DxUiTests.exe`

Tests the DirectX UI framework: controls, text input, rendering, theming, and accessibility.

| Family | File | Tests |
|--------|------|-------|
| TextInputBridge | `DxUiTests.TextInputBridge.cpp` | 171 |
| MultilineText | `DxUiTests.MultilineText.cpp` | 101 |
| Theme | `DxUiTests.Theme.cpp` | 54 |
| WindowHost | `DxUiTests.WindowHost.cpp` | 40 |
| TextField | `DxUiTests.TextField.cpp` | 39 |
| Grid | `DxUiTests.Grid.cpp` | 36 |
| ReadOnly | `DxUiTests.ReadOnly.cpp` | 25 |
| Animation | `DxUiTests.Animation.cpp` | 21 |
| Controls | `DxUiTests.Controls.cpp` | 18 |
| Tree | `DxUiTests.Tree.cpp` | 17 |
| Rendering | `DxUiTests.Rendering.cpp` | 15 |
| ComboBox | `DxUiTests.ComboBox.cpp` | 15 |
| Accessibility | `DxUiTests.Accessibility.cpp` | 10 |
| Tooltip | `DxUiTests.Tooltip.cpp` | 9 |

## 5. Performance Tests — 4 tests

**Project:** `Tests\PerformanceTests2\`  •  **Run:** `vstest.console.exe .\.build\x64\Debug\PerformanceTests2.dll`

CppUnitTest DLL for performance baselines.

| Test | Coverage |
|------|----------|
| FolderIconEnumerationPerfTest | Icon cache enumeration throughput |
| FolderIconEnumerationDuplicatePathPerfTest | Duplicate path icon edge cases |
| FolderViewRefreshDuplicatePathPerfTest | Folder view refresh optimisation |
| PerformanceTests2 (main) | FolderView enumeration access |

## 6. Viewer Plugin Tests — 17 tests

**Projects:** `Tests\ViewerPETests\` and `Tests\ViewerSqliteTests\`

| Project | Tests | Coverage |
|---------|-------|----------|
| ViewerPETests | 7 | PE, Web, ImgRaw, Text, Space, VLC viewers — DxUi combo host, long-run stability |
| ViewerSqliteTests | 10 | List tables, paged reads, sorting, DxUi host, scrolling, paging, theme, tab traversal |

## 7. Monitor / ETW Tests — 3 scenarios

**Project:** `Tests\MonitorTest\`  •  **Run:** `.\.build\x64\Debug\MonitorTest.exe`

Generates 150,000+ ETW trace messages across 3 burst scenarios to validate TraceLogging transport.

---

## Test Architecture

### Why Self-Tests Live in RedSalamander.exe

The self-test files are compiled **into** `RedSalamander.exe` when `ENABLE_TESTS` is defined because they are
**in-process integration tests** requiring direct access to:

- `extern FolderWindow g_folderWindow` — live window state
- `extern Common::Settings::Settings g_settings` — runtime settings
- Plugin manager, host services, UIAutomation — full application context
- FileOperations: async tick-driven state machine on the UI thread

`ENABLE_TESTS` is defined by default for first-party `Debug` and `ASan Debug` builds via
`Directory.Build.props`. Release builds stay production-clean unless a specific test-facing
project opts in explicitly.

### Conditional Coverage

Tests remain declared even when prerequisites are absent. Missing preconditions produce
`skipped` status (not removal):

- **Remote storage tests** — skip when connection profiles/secrets are absent
- **ReFS tests** — skip when no ReFS volume is available
- **Plugin tests** — skip when the required plugin is not loaded
- **Search service tests** — skip when the service is not running

### Self-Test Options

| Flag | Description |
|------|-------------|
| `--selftest` | Run all suites |
| `--commands-selftest` | Commands suite only |
| `--compare-selftest` | Compare Directories suite only |
| `--fileops-selftest` | File Operations suite only |
| `--selftest-fail-fast` | Abort after first failure |
| `--selftest-case=NAME` | Run a specific case (prefix match with trailing `_`) |
| `--selftest-timeout-multiplier=N` | Scale timeouts (>1.0 for slow machines) |

### Artifacts

Written to `%LOCALAPPDATA%\RedSalamander\SelfTest\last_run\`:
- `results.json` — per-case status, duration, failure reason
- `trace.txt` — diagnostic context log
- `perf_metrics.jsonl` — performance metric emissions

## Related Documentation

| Document | Purpose |
|----------|---------|
| [Testing_SelfTests.md](../Specs/Testing/Testing_SelfTests.md) | Result contract for self-test suites |
| [Testing_TestCoverage.md](../Specs/Testing/Testing_TestCoverage.md) | Per-suite test case listing |
| [Testing_PerformanceValidation.md](../Specs/Testing/Testing_PerformanceValidation.md) | Performance validation requirements |
| [Testing_SelfTestRemoteCredentials.md](../Specs/Testing/Testing_SelfTestRemoteCredentials.md) | Remote storage credential setup |
