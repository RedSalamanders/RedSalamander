# RedSalamander Test Infrastructure

## Overview

RedSalamander has multiple test surfaces covering UI automation, plugin
integration, file operations, DirectX rendering, performance baselines, tooling
scripts, and ETW diagnostics. Counts below are current as of 2026-05-15.

| Category | Suites | Tests | Framework |
|----------|--------|-------|-----------|
| Self-Tests (in-process) | 3 | 604 Commands listed cases, 149 Compare listed cases, 75 FileOps listed phases | Custom harnesses |
| DxUi Component Tests | 1 | 571 | Standalone harness |
| Performance Tests | 1 | 7 | CppUnitTest DLL |
| Viewer Plugin Tests | 2 | 17 | Standalone harness |
| File-System Plugin Tests | 1 | 5 | Standalone harness |
| Monitor/ETW Tests | 1 | 3 burst scenarios + 3 fast guards | Standalone harness |
| Tooling Script Tests | 1 folder + vcpkg scripts | 74 Pester-style/tool cases + 5 fast synthetic vcpkg merge cases | Pester / PowerShell |

Related specifications:
- `Specs/Testing/Testing_SelfTests.md` — result contract
- `Specs/Testing/Testing_TestCoverage.md` — per-case inventory
- `Specs/Testing/Testing_PerformanceValidation.md` — perf validation

## Quick Start

```powershell
.\Tools\Run-AllTests.ps1                      # Build Debug + run ALL self-tests
.\Tools\Run-AllTests.ps1 -Suite Full           # Build Debug + run self-tests, native tests, PerformanceTests2, and local script tests
.\Tools\Run-AllTests.ps1 -Suite Commands       # Single suite
.\Tools\Run-AllTests.ps1 -SkipBuild -FailFast  # Fast iteration
.\Tools\Run-AllTests.ps1 -TimeoutMultiplier 3  # Slow machines
.\Tools\Get-TestInventory.ps1 -Format Json     # Source-derived inventory manifest
.\.build\x64\Debug\RedSalamander.exe --selftest-list-cases  # Runner-native self-test case inventory
.\build.ps1 -ProjectName FileSystemCurlTests
.\.build\x64\Debug\FileSystemCurlTests.exe
```

---

## 1. Commands Self-Test Suite — 604 runner-listed cases

**Flag:** `--commands-selftest`
**Source:** `RedSalamander\SelfTest\Commands\Commands.SelfTest.cpp` + 12 `.cpp` family files
**Inventory:** `RedSalamander.exe --selftest-list-cases --commands-selftest`
lists 604 cases. The source fallback scan still reports 574 static
`SelfTest::RunCase` call sites because some helper call sites generate multiple
declared cases.

Tests UI automation, dialog interactions, preferences, shortcuts, themes, navigation,
and command dispatch inside the live application window.

| Family | File | Tests | Coverage |
|--------|------|-------|----------|
| Settings | `SelfTest\Commands\Commands.SelfTest.Settings.cpp` | 72 | Hot-reload, registry store, shortcut defaults |
| PluginConfig | `SelfTest\Commands\Commands.SelfTest.PluginConfig.cpp` | 17 | Plugin configuration, file-system plugin |
| Connections | `SelfTest\Commands\Commands.SelfTest.Connections.cpp` | 39 | Connection manager, credential prompts |
| Preferences Dispatch | `SelfTest\Commands\Commands.SelfTest.Preferences.Dispatch.cpp` | 169 | Preferences dialog — all categories, DxUi surfaces |
| CompareOptions | `SelfTest\Commands\Commands.SelfTest.CompareOptions.cpp` | 11 | Compare directories options, progress |
| Search | `SelfTest\Commands\Commands.SelfTest.Search.cpp` | 60 | Find dialog, local search index, quick search/filter |
| Shortcuts | `SelfTest\Commands\Commands.SelfTest.Shortcuts.cpp` | 1 | Shortcuts window grouped runner |
| ViewCommands | `SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp` | 59 | View commands, selection, sort, pane, tabs |
| FileOps | `SelfTest\Commands\Commands.SelfTest.FileOps.cpp` | 19 | File operations issues pane, speed limit |
| Navigation | `SelfTest\Commands\Commands.SelfTest.Navigation.cpp` | 46 | Navigation location, GoTo |
| Dialogs | `SelfTest\Commands\Commands.SelfTest.Dialogs.cpp` | 60 | About, fatal error, splash, change case, rename, filter, mask |
| ShellCommands | `SelfTest\Commands\Commands.SelfTest.ShellCommands.cpp` | 17 | Shell-integrated pane commands |

## 2. Compare Directories Self-Test Suite — 149 runner-listed cases

**Flag:** `--compare-selftest`
**Source:** `RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.cpp` + 3 included case files
**Inventory:** `RedSalamander.exe --selftest-list-cases --compare-selftest`
lists 149 cases. The source fallback scan still reports 141 static
`SelfTest::RunCase` call sites plus explicit setup/precondition result paths.

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

## 3. File Operations Self-Test Suite — 75 runner-listed phases

**Flag:** `--fileops-selftest`
**Source:** `RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.SelfTest.cpp` + 4 included phase files
**Inventory:** `RedSalamander.exe --selftest-list-cases --fileops-selftest`
lists 75 phases: setup, 73 active ordered phases, and cleanup.

Async tick-driven state machine testing file copy/move/delete operations end-to-end.
Organised into 12 families spanning phases 5–16.

| Family | Phases | Coverage |
|--------|--------|----------|
| Phase 05 — PreCalc | 7 | Pre-calculation settings, cancel, latency, mode switching |
| Phase 06 — PopupAndDelete | 5 | Popup, rate smoothing, bandwidth throttle, delete operations |
| Phase 07 — WatchAndParallelism | 17 | Watchers, cache, parallelism, concurrency |
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

## 5. Performance Tests — 7 tests

**Project:** `Tests\PerformanceTests2\`  •  **Run:** `vstest.console.exe .\.build\x64\Debug\PerformanceTests2.dll`

CppUnitTest DLL for performance baselines.

| Test | Coverage |
|------|----------|
| LargeFolderIconEnumeration_DuplicatePaths | Duplicate path icon edge cases |
| LargeFolderIconEnumeration_MixedItems | Icon cache enumeration throughput |
| FolderViewRefresh_PluginDuplicatePaths | Folder view refresh optimisation |
| FolderViewCompactMode_SetAppThemeCollapsesRowGapAndUpdatesHitTest | Compact mode hit-test/theme behavior |
| SplashScreenCloseGuardTriggersWhenCloseEventWasSignaled | Splash close guard |
| FileSystemPluginManagerInitializeFailsWhenNoPluginsAreDiscovered | File-system plugin manager empty discovery |
| ViewerPluginManagerInitializeFailsWhenNoPluginsAreDiscovered | Viewer plugin manager empty discovery |

## 6. Viewer Plugin Tests — 17 tests

**Projects:** `Tests\ViewerPETests\` and `Tests\ViewerSqliteTests\`

| Project | Tests | Coverage |
|---------|-------|----------|
| ViewerPETests | 7 | PE, Web, ImgRaw, Text, Space, VLC viewers — DxUi combo host, long-run stability |
| ViewerSqliteTests | 10 | List tables, paged reads, sorting, DxUi host, scrolling, paging, theme, tab traversal |

`ViewerPETests` runs most fresh-process viewer cases with the default
120-second child timeout. Its nested six-cycle shell-combo churn case uses a
dedicated 600-second parent timeout so valid long-run coverage is not killed
before the per-child viewer checks can finish and report their own results.

## 7. File-System Plugin Tests — 7 tests

**Project:** `Tests\FileSystemCurlTests\`  •  **Run:** `.\.build\x64\Debug\FileSystemCurlTests.exe`

Focused deterministic coverage for `Plugins\FileSystemCurl\` helpers that do not require live remote credentials.

| Family | Tests | Coverage |
|--------|-------|----------|
| IMAP naming and UID parsing | 3 | Preferred `<subject> [uid].eml`, direct `<uid>.eml`, malformed-name rejection including retired decorated names |
| IMAP subject decoding | 1 | RFC2047 Q/B decoding, mixed plain/encoded fragments, UTF-8 emoji, non-UTF code pages, malformed sanitized fragments |
| IMAP mailbox status parsing | 1 | `STATUS` counts: messages, recent, uidNext, uidValidity, unseen |
| IMAP properties perf model | 1 | Command-count guard proving single-message Properties stays constant with mailbox size |
| IMAP listing metadata repair | 1 | Batch plan guard proving large missing summary sets are retried instead of skipped |

The executable also supports `--perf` for a lightweight deterministic probe of IMAP leaf parsing, subject decoding, leaf building, `STATUS` parsing, repair batch planning, and the message Properties command-count model.

## 8. Monitor / ETW Tests — 3 burst scenarios plus fast guards

**Project:** `Tests\MonitorTest\`  •  **Run:** `.\.build\x64\Debug\MonitorTest.exe`

Generates 150,000+ ETW trace messages across 3 burst scenarios to validate TraceLogging transport.
Fast targeted guards include `--diagnostics-gate-selftest`, `--scrollbar-model-selftest`, and `--document-model-selftest`.

## 9. Tooling Script Tests

**Run locally/full:** `Invoke-Pester .\Tools\Tests`

**Run in artifact-only CI jobs:** `Invoke-Pester .\Tools\Tests -ExcludeTag RequiresBuildToolchain`

| File | Cases | Coverage |
|------|-------|----------|
| `BuildProjectSelection.Tests.ps1` | 4 | Project selection and direct vcxproj builds |
| `MSBuildInvocation.Tests.ps1` | 8 | MSBuild invocation planning and diagnostic parsing |
| `ProcessStreaming.Tests.ps1` | 2 | Process output streaming and logging |
| `RedSalamanderPluginDeployment.Tests.ps1` | 1 | Targeted RedSalamander build repopulates sibling binaries/plugins and plugin language resources; tagged `RequiresBuildToolchain`, bounded, and logged |
| `ResourceLocalizationContracts.Tests.ps1` | 1 | Resource placeholder positional-order and satellite placeholder-equivalence contract |
| `RunAllTestsPlan.Tests.ps1` | 8 | Full runner test-plan enumeration, aggregate artifact, and result-coverage validation |
| `SanitizedEnvironment.Tests.ps1` | 2 | Child process environment normalization |
| `TestHarnessSourceContracts.Tests.ps1` | 14 | Source guards for test harness CLI/error handling, case-listing, result-emission, duplicate-name contracts, CompareDirectories listed-case coverage, async file-operations self-test prompts, and FileOperations prefix filters |
| `TestInventory.Tests.ps1` | 5 | Source-derived test inventory manifest, FileOperations phase-order drift guard, and doc-count lint |
| `VcpkgInstallSafety.Tests.ps1` | 5 | vcpkg path/triplet safety |
| `Versioning.Tests.ps1` | 4 | Local build-number reuse/allocation |
| `WingetValidation.Tests.ps1` | 11 | Winget validation warning suppression, failure propagation, portable manifest metadata, and VC runtime ZIP helper coverage |

Fast vcpkg merge coverage:

- `.\Tests\vcpkg-merge-synthetic-test.ps1` — 5 synthetic lock/merge tests.

Manual-only intrusive validation:

- `.\Tests\vcpkg-merge-lock-validation.ps1` — 3 vcpkg install/lock validation
  scenarios; this is intentionally not part of `Run-AllTests.ps1 -Suite Full`
  or PR CI because it runs vcpkg install flows and mutates `.build`.

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
| `--selftest-timeout-multiplier=N` | Scale timeouts by a finite value clamped to `[0.1, 100.0]`; invalid values fail fast |

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
