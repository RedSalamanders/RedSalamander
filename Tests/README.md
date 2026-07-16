# RedSalamander Test Infrastructure

## Overview

RedSalamander has multiple test surfaces covering UI automation, plugin
integration, file operations, DirectX rendering, performance baselines, tooling
scripts, and ETW diagnostics. Counts below are current as of 2026-07-12.

| Category | Suites | Tests | Framework |
|----------|--------|-------|-----------|
| Self-Tests (in-process) | 3 | 809 Commands listed cases, 256 Compare listed cases, 128 FileOps listed phases | Custom harnesses |
| DxUi Component Tests | 1 | 637 | Standalone harness |
| Performance Tests | 1 | 14 | CppUnitTest DLL |
| Viewer Plugin Tests | 2 | 17 | Standalone harness |
| File-System Plugin Tests | 1 | 8 | Standalone harness |
| Monitor/ETW Tests | 1 | 3 burst scenarios + 3 fast guards | Standalone harness |
| RedConfigure Tests | 1 | 23 | Standalone harness |
| Tooling Script Tests | 1 folder + vcpkg scripts | 294 Pester-style/tool cases + 5 fast synthetic vcpkg merge cases | Pester / PowerShell |

Related specifications:
- `Specs/Testing/Testing_SelfTests.md` — result contract
- `Specs/Testing/Testing_TestCoverage.md` — per-case inventory
- `Specs/Testing/Testing_PerformanceValidation.md` — perf validation

## Quick Start

```powershell
.\Tools\Run-AllTests.ps1                      # Build Debug + run ALL self-tests
.\Tools\Run-AllTests.ps1 -Suite CI             # Build Debug + run the GitHub Actions PR gate locally
.\Tools\Run-AllTests.ps1 -Suite Full           # Build Debug + run self-tests, native tests, PerformanceTests2, and local script tests
.\Tools\Run-AllTests.ps1 -Suite Commands       # Single suite
.\Tools\Run-AllTests.ps1 -SkipBuild -FailFast  # Fast iteration
.\Tools\Run-AllTests.ps1 -TimeoutMultiplier 3  # Slow machines
.\Tools\Get-TestInventory.ps1 -Format Json     # Source-derived inventory manifest
.\.build\x64\Debug\RedSalamander.exe --selftest-list-cases  # Runner-native self-test case inventory
.\build.ps1 -ProjectName FileSystemCurlTests
.\.build\x64\Debug\FileSystemCurlTests.exe
.\build.ps1 -ProjectName RedConfigureTests
.\.build\x64\Debug\RedConfigureTests.exe
```

---

## 1. Commands Self-Test Suite — 809 runner-listed cases

**Flag:** `--commands-selftest`
**Source:** `RedSalamander\SelfTest\Commands\Commands.SelfTest.cpp` + 13 `.cpp` family files
**Inventory:** `RedSalamander.exe --selftest-list-cases --commands-selftest`
lists 809 cases. The source fallback scan reports 707 static
`SelfTest::RunCase` call sites: 706 family registrations plus the orchestrator's
single pre-dispatch isolation-failure registration. Some helper call sites also
generate multiple declared cases.

Tests UI automation, dialog interactions, preferences, shortcuts, themes, navigation,
and command dispatch inside the live application window.

| Family | File | Tests | Coverage |
|--------|------|-------|----------|
| Settings | `SelfTest\Commands\Commands.SelfTest.Settings.cpp` | 99 | Hot-reload, section-scoped/lossless SettingsStore recovery, future-schema save blocking, Windows session-end durability, shortcut defaults, bounded UIA dispatch, embedded-preview HWND ownership/cardinality |
| BatchRename | `SelfTest\Commands\Commands.SelfTest.BatchRename.cpp` | 1 | Batch Rename dialog source guard |
| PluginConfig | `SelfTest\Commands\Commands.SelfTest.PluginConfig.cpp` | 19 | Plugin configuration, file-system plugin |
| Connections | `SelfTest\Commands\Commands.SelfTest.Connections.cpp` | 44 | Connection manager, credential prompts, plugin-backed MTP picker, session-first credential persistence failure |
| Preferences Dispatch | `SelfTest\Commands\Commands.SelfTest.Preferences.Dispatch.cpp` | 173 | Preferences dialog — all categories, DxUi surfaces |
| CompareOptions | `SelfTest\Commands\Commands.SelfTest.CompareOptions.cpp` | 15 | Compare directories options, progress |
| Search | `SelfTest\Commands\Commands.SelfTest.Search.cpp` | 71 | Find dialog, local search index, quick search/filter |
| Shortcuts | `SelfTest\Commands\Commands.SelfTest.Shortcuts.cpp` | 1 | Shortcuts window grouped runner |
| ViewCommands | `SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp` | 119 | View commands, selection, sort, pane, tabs, FolderView rendering-alert persistence, draw-item brush reuse, refresh-to-paint telemetry, IconCache live-failure retry, DPI repaint |
| FileOps | `SelfTest\Commands\Commands.SelfTest.FileOps.cpp` | 28 | File operations issues pane, speed limit |
| Navigation | `SelfTest\Commands\Commands.SelfTest.Navigation.cpp` | 47 | Navigation location, GoTo |
| Dialogs | `SelfTest\Commands\Commands.SelfTest.Dialogs.cpp` | 61 | About, fatal error, splash, change case, rename, filter, mask |
| ShellCommands | `SelfTest\Commands\Commands.SelfTest.ShellCommands.cpp` | 28 | Shell-integrated pane commands |

## 2. Compare Directories Self-Test Suite — 256 runner-listed cases

**Flag:** `--compare-selftest`
**Source:** `RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.cpp` + 4 included case files
**Inventory:** `RedSalamander.exe --selftest-list-cases --compare-selftest`
lists 256 cases. The source fallback scan reports 249 static
`SelfTest::RunCase` call sites plus explicit setup/precondition result paths.

Tests the compare-directories engine including local/remote search, indexing, and session logic.

| Family | Tests | Coverage |
|--------|-------|----------|
| Core session | ~25 | Unique, size, content, unicode, subdirs, ignore patterns, invalidation |
| Search service | 39 | CLI, bootstrap, cold start, compact, maintenance, multi-client, deleted-root rebuild purge, candidate-authorization warning, transient authorization cache, external SQLite rotation |
| Local search | 18 | QI, callbacks, backend prefs, tree walking, wildcards, content |
| Local index core | 12 | Snapshot reload, journal replay, corruption rebuild, SQLite |
| SQLite index store | 6 | Schema bootstrap, compaction, checkpoint, upgrade |
| Host fallback | 7 | Plugin path root, degraded IO, cancel, access denied |
| Remote filesystems | 9 | S3, OneDrive, SharePoint, FTP directories and size callbacks |
| MTP/PTP | 52 | Fake/live MTP contract, shared identity helpers, streaming reader, overwrite safety, watchdog/quarantine, worker reuse, WPD session/path cache |
| Google Drive | 3 | Plugin contract, client ID, refresh token |
| OAuth / credentials | 3 | Token storage, auth mode, Windows Hello cache |
| Directory size | 3 | Local, dummy, 7z filesystem callbacks |
| Search text helpers | 2 | Text matching, decoding |
| Misc (concurrency, caching, UI) | ~32 | Crash quarantine, setCompareEnabled, uiVersion, etc. |

## 3. File Operations Self-Test Suite — 128 runner-listed phases

**Flag:** `--fileops-selftest`
**Source:** `RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.SelfTest.cpp` + 4 included phase files
**Inventory:** `RedSalamander.exe --selftest-list-cases --fileops-selftest`
lists 128 phases: setup, 126 active ordered phases, and cleanup.

Async tick-driven state machine testing file copy/move/delete operations end-to-end.
Organised into 12 families spanning phases 5–16.

| Family | Phases | Coverage |
|--------|--------|----------|
| Phase 05 — PreCalc | 7 | Pre-calculation settings, cancel, latency, mode switching, preflight Speed Limit affordance |
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

## 4. DxUi Component Tests — 637 tests

**Project:** `Tests\DxUiTests\`  •  **Run:** `.\.build\x64\Debug\DxUiTests.exe`

Tests the DirectX UI framework: controls, text input, rendering, theming, and accessibility.
Run HWND focus-sensitive suites such as `NativeTextInput` serially when collecting closeout evidence; they create real test windows and can legitimately affect process/global Win32 focus. Real Win32 focus/caret/foreground assertions must use `TryFocusDxUiTestWindow` or `TryActivateDxUiTestWindow`, and must emit an explicit `SKIPPED:` reason when the current desktop session cannot provide the required capability.

| Family | File | Tests |
|--------|------|-------|
| MultilineText | `DxUiTests.MultilineText.cpp` | 106 |
| Theme | `DxUiTests.Theme.cpp` | 76 |
| WindowHost | `DxUiTests.WindowHost.cpp` | 52 |
| TextField | `DxUiTests.TextField.cpp` | 62 |
| Grid | `DxUiTests.Grid.cpp` | 49 |
| ReadOnly | `DxUiTests.ReadOnly.cpp` | 24 |
| Animation | `DxUiTests.Animation.cpp` | 24 |
| Controls | `DxUiTests.Controls.cpp` | 22 |
| Tree | `DxUiTests.Tree.cpp` | 23 |
| Rendering | `DxUiTests.Rendering.cpp` | 22 |
| ComboBox | `DxUiTests.ComboBox.cpp` | 24 |
| Accessibility | `DxUiTests.Accessibility.cpp` | 32 |
| Tooltip | `DxUiTests.Tooltip.cpp` | 10 |
| NativeTextInput | `DxUiTests.NativeTextInput.cpp` | 118 |

## 5. Performance Tests — 14 tests

**Project:** `Tests\PerformanceTests2\`  •  **Run:** `vstest.console.exe .\.build\x64\Debug\PerformanceTests2.dll`

CppUnitTest DLL for performance baselines.

| Test | Coverage |
|------|----------|
| LargeFolderIconEnumeration_DuplicatePaths | Duplicate path icon edge cases |
| LargeFolderIconEnumeration_MixedItems | Icon cache enumeration throughput |
| FolderViewRefresh_PluginDuplicatePaths | Folder view refresh optimisation |
| FolderViewCompactMode_SetAppThemeCollapsesRowGapAndUpdatesHitTest | Compact mode hit-test/theme behavior |
| VisibleColumnWidths_DifferentColumnsDoNotShareGlobalMax | Variable column widths stay column-local |
| VisibleColumnWidths_DetailedAndMetadataLinesStayColumnLocal | Detailed/metadata width calculations stay column-local |
| ScrollStops_FirstRightSkipsInitialLeftGap | First-column horizontal scroll snapping |
| ScrollStops_FirstLeftRestoresInitialLeftGap | Left-leading gap hit-test and snap-back behavior |
| SortPolicy_ParallelPathStartsAtLargeFolderThreshold | FolderView parallel-sort threshold policy |
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

## 7. File-System Plugin Tests — 8 tests

**Project:** `Tests\FileSystemCurlTests\`  •  **Run:** `.\.build\x64\Debug\FileSystemCurlTests.exe`

Focused deterministic coverage for `Plugins\FileSystemCurl\` helpers that do not require live remote credentials.

| Family | Tests | Coverage |
|--------|-------|----------|
| IMAP naming and UID parsing | 3 | Preferred `<subject> [uid].eml`, direct `<uid>.eml`, malformed-name rejection including retired decorated names |
| IMAP subject decoding | 1 | RFC2047 Q/B decoding, mixed plain/encoded fragments, UTF-8 emoji, non-UTF code pages, malformed sanitized fragments |
| IMAP mailbox status parsing | 1 | `STATUS` counts: messages, recent, uidNext, uidValidity, unseen |
| IMAP properties perf model | 1 | Command-count guard proving single-message Properties stays constant with mailbox size |
| IMAP listing metadata repair | 2 | Batch plan guard proving large missing summary sets are retried instead of skipped, plus bounded per-listing repair fetch budget coverage |

The executable also supports `--perf` for a lightweight deterministic probe of IMAP leaf parsing, subject decoding, leaf building, `STATUS` parsing, repair batch planning, and the message Properties command-count model.

## 8. RedConfigure Tests — 23 tests

**Project:** `Tests\RedConfigureTests\`  •  **Run:** `.\.build\x64\Debug\RedConfigureTests.exe`

Focused deterministic coverage for RedConfigure page definitions, workspace discovery, theme JSON5 parsing/export/validation, RC parsing/writing, placeholder validation, translation table search/filter/sort, theme preview models, and session export/loading behavior.

## 9. Monitor / ETW Tests — 3 burst scenarios plus fast guards

**Project:** `Tests\MonitorTest\`  •  **Run:** `.\.build\x64\Debug\MonitorTest.exe`

Generates 150,000+ ETW trace messages across 3 burst scenarios to validate TraceLogging transport.
Fast targeted guards include `--diagnostics-gate-selftest`, `--scrollbar-model-selftest`, and `--document-model-selftest`.

## 10. Tooling Script Tests

**Run locally/full:** `Invoke-Pester .\Tools\Tests`

**Run in artifact-only CI jobs:** `.\Tools\Run-AllTests.ps1 -Suite CI -SkipBuild`

Suite CI includes the deterministic PluginContractTests, SettingsSchemaTests, and CrashHandlingTests
executables. The pull-request workflow separately performs a Debug ARM64 build-only gate; hosted x64
runners do not execute the ARM64 artifacts.

| File | Cases | Coverage |
|------|-------|----------|
| `BuildOutputProcess.Tests.ps1` | 15 | Build preflight self-test protection plus creation-time job containment, inherited-handle allowlisting, Unicode/stream fidelity, launch-failure cleanup, direct-child delegation, parallel-runspace and stale-descendant exclusion, owner diagnostics, abandoned-owner contamination, and residual compiler diagnostics |
| `BuildProjectSelection.Tests.ps1` | 4 | Project selection and direct vcxproj builds |
| `DocumentationDriftContracts.Tests.ps1` | 9 | Source/spec documentation drift guards, including split File Operations popup coverage and archived remediation status |
| `HwndRenderTargetResourcesSourceContracts.Tests.ps1` | 3 | Shared HWND render-target/brush lifecycle with FunctionBar and StatusBar policy, invalidation, and teardown retained locally |
| `ModalWindowShellSourceContracts.Tests.ps1` | 3 | Shared About/Fatal Error modal owner, message-loop, quit-propagation, and owner-restoration boundaries |
| `MSBuildInvocation.Tests.ps1` | 8 | MSBuild invocation planning and diagnostic parsing |
| `MtpLiveCloseout.Tests.ps1` | 4 | MTP live closeout wrapper safety, archival, and environment-restoration contracts |
| `PackedFileInfoBufferSourceContracts.Tests.ps1` | 3 | Checked packed FileInfo sizing/construction/traversal shared by six buffered COM facades while streaming and fixture variants remain local |
| `PluginConfigurationConsolidationSourceContracts.Tests.ps1` | 3 | Common plugin-configuration schema/codec reuse and lossless Preferences commits |
| `PluginLifetimeConsolidationSourceContracts.Tests.ps1` | 3 | Shared callback generation/drain state, callback-return module-pin transfer, and case-insensitive manager lookup leases |
| `PostedPayloadCoalescingSourceContracts.Tests.ps1` | 3 | Queue-head-safe keyed payload coalescing for Compare progress and Find result/progress drains |
| `ProcessStreaming.Tests.ps1` | 4 | Process output streaming, logging, and kill-on-close descendant containment |
| `RedSalamanderPluginDeployment.Tests.ps1` | 1 | Targeted RedSalamander build repopulates sibling binaries/plugins and plugin language resources; tagged `RequiresBuildToolchain`, bounded, and logged |
| `ResourceLocalizationContracts.Tests.ps1` | 5 | Resource placeholder positional-order, satellite placeholder-equivalence, and language-neutral embedded-only string contracts |
| `RunAllTestsPlan.Tests.ps1` | 33 | CI/Full runner test-plan enumeration, unified test-sandbox root selection, Pester 3/newer invocation compatibility, dead-PID stale run cleanup, disk-audit evidence, and legacy cleanup target planning including locked-target failure reporting, optional native perf-budget gate forwarding, repeat/shuffle and injected classifier-proof argument forwarding, focused-filter shuffle-triage planning, invariant timeout formatting, blocking failure classification with shuffle-triage evidence, reviewed quarantine validation and repair-lane planning, GitHub step summary formatting, aggregate artifact, per-case history/dashboard output, and result-coverage validation |
| `SanitizedEnvironment.Tests.ps1` | 5 | Child process environment normalization |
| `ShowPerfRuns.Tests.ps1` | 7 | Perf-run report parsing and filtering |
| `TestHarnessSourceContracts.Tests.ps1` | 128 | Source guards for test harness CLI/error handling, artifact-operation serialization and creation-time Job Object/handle-list containment, case-listing, repeat/shuffle controls, injected classifier-proof hooks, explicit-order Commands case isolation, shared TestSupport sandbox/environment policy, bounded viewer message-pump/snapshot polling, contained concurrent child-process execution, native unified TestSandbox root/run-id consumption, native TestSandbox scratch acquisition, FileOperations alternate-volume TestSandbox scratch routing/pruning, RedConfigureTests, ViewerSqliteTests, CrashHandlingTests unified TestSandbox scratch routing, Commands plugin-config native TestSandbox scratch routing, ShellCommands shortcut-save native TestSandbox routing, PerformanceTests2 unified TestSandbox scratch routing, Commands BatchRename window fixture native TestSandbox routing, ViewerPETests unified TestSandbox fixture routing, DxUiTests generated artifact default TestSandbox routing, MTP fake journal `LOCALAPPDATA` TestSandbox routing, Compare dummy filesystem native TestSandbox scratch routing, Compare foreground service sandboxed stdout capture and kill-on-close JobObject isolation, broad raw temp/profile/legacy-root source guard, local index snapshot reload advisory timing, raw self-test wait scaling through `SelfTest::ScaleTimeout`, legacy TestSandbox cleanup script safety, failed-status warning behavior, and runner invocation, nightly shuffle-repeat workflow isolation, CompareDirectories and FileOperations explicit-order seeded shuffle, FileOperations native repeat aggregation, result-emission, duplicate-name contracts, CompareDirectories listed-case coverage, self-test fatal-modal bypass, partial crash-result preservation, async file-operations self-test prompts, FileOperations prefix filters, FileOperations pause-point centralization, FileOperations bridge IO decorator source-size fault seam, FileOperations bridge create-directory race env-helper reuse, FileOperations issues-pane focus-restore helper reuse, FolderView owned threadpool submit helper reuse, FolderView thumbnail stat helper reuse, FolderView pending-to-paint metric shape reuse, IconCache path failure-store duplicate-race contracts, MTP RAII owner/directory-info handoff contracts, LocalSearch snapshot temp-path exception logging contracts, SearchAndIndex callback exception logging contracts, DxUi native text-input bounded formatting contracts, DxUi focus-sensitive Win32 assertion desktop-probe contracts, FileSystemMtp pragma-warning rationale contracts, shared ordinal string helper usage, shared truthy env-flag helper usage for FolderView WARP/perf flags, shared ViewerSpace opt-in env-flag parser usage, shared DxUi modal loop/quit propagation for archive prompts, FolderView perf-budget strict-mode contracts, FolderView overlay perf advisory sample-sufficiency contract, Riptide/Floodgate source contracts, multi-plugin metadata/localization ownership, ViewerPE exact-reader terminal delivery, ViewerSpace non-waiting close/unload gating, Quick Search no-match focus reactivation, native-menu cascade reacquisition, pane-restoration enumeration/focus handoff, and menu-bar keyboard-open stale-pointer guard handling |
| `TestInventory.Tests.ps1` | 5 | Source-derived test inventory manifest, exhaustive FileOperationsSettings field-case coverage, FileOperations phase-order drift guard, and exhaustive doc-count lint |
| `TestRunArchive.Tests.ps1` | 6 | TestRun archive size, machine-profile, curated-evidence, and empty-path policy guards |
| `ThemeDistributionContracts.Tests.ps1` | 9 | Theme schema, package-assets, licensing, and distribution contract guards |
| `VcpkgInstallSafety.Tests.ps1` | 5 | vcpkg path/triplet safety |
| `VerifyNoProductionGetCursorPos.Tests.ps1` | 1 | Production sources avoid `GetCursorPos` outside annotated diagnostics |
| `Versioning.Tests.ps1` | 4 | Local build-number reuse/allocation |
| `ViewerChromeSourceContracts.Tests.ps1` | 6 | Viewer chrome keyboard routing, launcher subsystem, and Escape contract guards |
| `WingetValidation.Tests.ps1` | 17 | Winget validation warning suppression, failure propagation, portable manifest metadata, and VC runtime ZIP helper coverage |

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
| `--selftest-repeat=N` | Run each matched case N times in-process; Commands/Compare use `RunCase`, FileOperations expands phase-state runs, and result rows include `repeat_index` |
| `--selftest-shuffle=SEED` | Commands and CompareDirectories seeded case-order shuffle; FileOperations uses seeded phase-order shuffle while preserving setup/cleanup rows |
| `--selftest-crash-case=NAME` | Debug-only crash-signal proof hook; raises an access violation when the exact case starts |
| `--selftest-flaky-proof-case=NAME` | Debug-only classifier proof hook; the named case fails in suite context, then passes exact-case and shuffle-triage retries |
| `--selftest-order-proof-case=NAME` | Debug-only classifier proof hook; the named case fails in suite/shuffle context, then passes exact-case retry so shuffle triage classifies it as an isolation/order regression |
| `--selftest-list-cases` | Emit runner-native case inventory without overwriting prior run artifacts |
| `--selftest-timeout-multiplier=N` | Scale timeouts by a finite value clamped to `[0.1, 100.0]`; invalid values fail fast |

PowerShell harnesses that need the exit code or final artifacts from `RedSalamander.exe`
self-test runs must launch the GUI-subsystem executable with `Start-Process -Wait -PassThru`
or `System.Diagnostics.Process`. Direct invocation can return before `results.json` and
the repository archive are finalized.

### Shared native test support

Standalone first-party harnesses use `TestSupport/TestSupport.h` for scoped
environment changes and TestSandbox directory acquisition. Harness adapters must
state their policy choices explicitly: scratch versus artifacts, leaf inclusion,
empty-leaf fallback, and clean-versus-retain behavior. UI waits use the bounded
message-pump and typed-snapshot helpers so timeout diagnostics include the
operation, budget, elapsed time, and dispatched-message count.

Captured child execution uses `TestSupport/ChildProcess.h`. Callers pass an
executable plus structured arguments; the runner owns all process, thread, pipe,
and JobObject handles through WIL, starts the root suspended, allowlists inherited
standard handles, assigns the kill-on-close job before resume, and drains stdout
and stderr concurrently even after their retained prefixes reach their configured
limits. Timeout and cancellation contain the whole process tree. A nonzero exit is
a completed launch whose exit code and separate streams remain available to the
test oracle.

### Artifacts

When launched through `Tools\Run-AllTests.ps1`, artifacts are written under
`REDSALAMANDER_TEST_ROOT\runs\<runId>\artifacts\selftest\last_run\`; by default
the runner sets `REDSALAMANDER_TEST_ROOT` to `.build\TestSandbox`, sets
`REDSALAMANDER_TEST_RUN_ID`, ignores/clears inherited `REDSALAMANDER_SELFTEST_ROOT`, and
native self-tests resolve this path directly.
Before child tests launch, the runner applies `Tools\Clean-TestSandbox.ps1` to remove known legacy
self-test and temp roots. Running that script directly is dry-run by default; pass `-Apply` only when
you intend to delete the listed targets. Locked or access-denied legacy targets are reported as
warning-backed failed cleanup rows and must not prevent the requested suite from starting.
Tooling Pester tests that create scratch files or synthetic run trees use
`Tools\TestRunPlan.ps1` `New-RSTestSandboxScratchDirectory(...)` and write under
`REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\tools-pester\<case>\` instead of
the process temp directory.

- `results.json` — per-case status, duration, failure reason
- `trace.txt` — diagnostic context log
- `perf_metrics.jsonl` — performance metric emissions
- `run-all-tests-results.json` — runner-owned aggregate summary with `test_root`, `run_id`, non-blocking `test_sandbox_audit`, blocking classification counts, retry/shuffle-triage evidence, quarantine metadata, repair-lane attempt evidence, and GitHub step-summary source data
- `run-all-tests-case-history.jsonl` — per-case and retry rows keyed by harness/case with duration, status, reason, classification, seed, and attempt
- `run-all-tests-dashboard.md` — compact per-run slow-case and failure/retry dashboard derived from the history rows

## Related Documentation

| Document | Purpose |
|----------|---------|
| [Testing_SelfTests.md](../Specs/Testing/Testing_SelfTests.md) | Result contract for self-test suites |
| [Testing_TestCoverage.md](../Specs/Testing/Testing_TestCoverage.md) | Per-suite test case listing |
| [Testing_PerformanceValidation.md](../Specs/Testing/Testing_PerformanceValidation.md) | Performance validation requirements |
| [Testing_SelfTestRemoteCredentials.md](../Specs/Testing/Testing_SelfTestRemoteCredentials.md) | Remote storage credential setup |
