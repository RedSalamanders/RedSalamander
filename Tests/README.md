# RedSalamander Test Infrastructure

## Overview

RedSalamander has multiple test surfaces covering UI automation, plugin
integration, file operations, DirectX rendering, performance baselines, tooling
scripts, and ETW diagnostics. Mutable counts are deliberately not checked into
this guide. Run `Tools\Get-TestInventory.ps1 -Format Json` for the current
source/run-plan manifest and `RedSalamander.exe --selftest-list-cases` for the
current in-product case list.

| Category | Canonical execution kind | Inventory authority |
|----------|--------------------------|---------------------|
| In-product self-tests | `SelfTest` | Runner-native case listing plus `Tools\TestRunPlan.ps1` |
| Native component/plugin tests | `Executable` | `Tests\*.vcxproj` reconciled with CI/Full run plans |
| Performance tests | `CppUnitTest` | `PerformanceTests2.vcxproj` plus CI/Full run plans |
| Tooling tests | `Pester` / `PowerShellScript` | CI/Full run plans plus source-derived manifest |

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

## 1. Commands Self-Test Suite

**Flag:** `--commands-selftest`
**Source:** `RedSalamander\SelfTest\Commands\Commands.SelfTest.cpp` + 13 `.cpp` family files
**Inventory:** `RedSalamander.exe --selftest-list-cases --commands-selftest`
emits the live case list. `Tools\Get-TestInventory.ps1 -Format Json` derives the
static family/coordinator registration breakdown. Some helper call sites generate
multiple declared cases, so the two inventories intentionally measure different
things.

Tests UI automation, dialog interactions, preferences, shortcuts, themes, navigation,
and command dispatch inside the live application window.

| Family | File | Coverage |
|--------|------|----------|
| Settings | `SelfTest\Commands\Commands.SelfTest.Settings.cpp` | Hot-reload, section-scoped/lossless SettingsStore recovery, future-schema save blocking, Windows session-end durability, shortcut defaults, bounded UIA dispatch, embedded-preview HWND ownership/cardinality |
| BatchRename | `SelfTest\Commands\Commands.SelfTest.BatchRename.cpp` | Batch Rename dialog source guard |
| PluginConfig | `SelfTest\Commands\Commands.SelfTest.PluginConfig.cpp` | Plugin configuration, file-system plugin |
| Connections | `SelfTest\Commands\Commands.SelfTest.Connections.cpp` | Connection manager, credential prompts, plugin-backed MTP picker, session-first credential persistence failure |
| Preferences Dispatch | `SelfTest\Commands\Commands.SelfTest.Preferences.Dispatch.cpp` | Preferences dialog — all categories, DxUi surfaces |
| CompareOptions | `SelfTest\Commands\Commands.SelfTest.CompareOptions.cpp` | Compare directories options, progress |
| Search | `SelfTest\Commands\Commands.SelfTest.Search.cpp` | Find dialog, local search index, quick search/filter |
| Shortcuts | `SelfTest\Commands\Commands.SelfTest.Shortcuts.cpp` | Shortcuts window grouped runner |
| ViewCommands | `SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp` | View commands, selection, sort, pane, tabs, FolderView rendering-alert persistence, draw-item brush reuse, refresh-to-paint telemetry, IconCache live-failure retry, DPI repaint |
| FileOps | `SelfTest\Commands\Commands.SelfTest.FileOps.cpp` | File operations issues pane, speed limit |
| Navigation | `SelfTest\Commands\Commands.SelfTest.Navigation.cpp` | Navigation location, GoTo |
| Dialogs | `SelfTest\Commands\Commands.SelfTest.Dialogs.cpp` | About, fatal error, splash, change case, rename, filter, mask |
| ShellCommands | `SelfTest\Commands\Commands.SelfTest.ShellCommands.cpp` | Shell-integrated pane commands, bounded FolderView drop parsing/targeting/effect semantics, pointer target integrity, and verified move-clipboard invalidation |

## 2. Compare Directories Self-Test Suite

**Flag:** `--compare-selftest`
**Source:** `RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.cpp` + 4 included case files
**Inventory:** `RedSalamander.exe --selftest-list-cases --compare-selftest`
emits the live case list; the source-derived manifest reports static registrations
and keeps explicit setup/precondition result paths distinct.

Tests the compare-directories engine including local/remote search, indexing, and session logic.

| Family | Coverage |
|--------|----------|
| Core session | Unique, size, content, unicode, subdirs, ignore patterns, invalidation |
| Search service | CLI, bootstrap, cold start, compact, maintenance, multi-client, deleted-root rebuild purge, candidate-authorization warning, transient authorization cache, external SQLite rotation |
| Local search | QI, callbacks, backend prefs, tree walking, wildcards, content |
| Local index core | Snapshot reload, journal replay, corruption rebuild, SQLite |
| SQLite index store | Schema bootstrap, compaction, checkpoint, upgrade |
| Host fallback | Plugin path root, degraded IO, cancel, access denied |
| Remote filesystems | S3, OneDrive, SharePoint, FTP directories and size callbacks |
| MTP/PTP | Fake/live MTP contract, shared identity helpers, streaming reader, overwrite safety, watchdog/quarantine, worker reuse, WPD session/path cache |
| Google Drive | Plugin contract, client ID, refresh token |
| OAuth / credentials | Token storage, auth mode, Windows Hello cache |
| Directory size | Local, dummy, 7z filesystem callbacks |
| Search text helpers | Text matching, decoding |
| Misc (concurrency, caching, UI) | Crash quarantine, setCompareEnabled, uiVersion, etc. |

## 3. File Operations Self-Test Suite

**Flag:** `--fileops-selftest`
**Source:** `RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.SelfTest.cpp` + 4 included phase files
**Inventory:** `RedSalamander.exe --selftest-list-cases --fileops-selftest`
emits setup, the active ordered phases, and cleanup. The source-derived manifest
validates the phase order against the `Step` enum without freezing a total here.

Async tick-driven state machine testing file copy/move/delete operations end-to-end.
Organised into 12 families spanning phases 5–16.

| Family | Coverage |
|--------|----------|
| Phase 05 — PreCalc | Pre-calculation settings, cancel, latency, mode switching, preflight Speed Limit affordance |
| Phase 06 — PopupAndDelete | Popup, rate smoothing, bandwidth throttle, delete operations |
| Phase 07 — WatchAndParallelism | Watchers, cache, parallelism, concurrency |
| Phase 08 — Validation | Defaults, destinations, size bytes |
| Phase 09 — ConflictPrompt | Overwrite, apply-to-all, skip, retry |
| Phase 10 — DeleteValidation | Permanent delete |
| Phase 11 — BridgeAndConnections | Cross-filesystem bridges, connection overrides |
| Phase 12 — Reparse | Reparse point policy |
| Phase 13 — PostMortem | Post-mortem diagnostics |
| Phase 14 — PopupLifetime | Popup lifetime guard |
| Phase 15 — FileSystem7z | 7z filesystem operations |
| Phase 16 — Remote | FTP, SFTP, SCP, IMAP, S3, OneDrive, SharePoint |

---

## 4. DxUi Component Tests

**Project:** `Tests\DxUiTests\`  •  **Run:** `.\.build\x64\Debug\DxUiTests.exe`

Tests the DirectX UI framework: controls, text input, rendering, theming, and accessibility.
Run HWND focus-sensitive suites such as `NativeTextInput` serially when collecting closeout evidence; they create real test windows and can legitimately affect process/global Win32 focus. Real Win32 focus/caret/foreground assertions must use `TryFocusDxUiTestWindow` or `TryActivateDxUiTestWindow`, and must emit an explicit `SKIPPED:` reason when the current desktop session cannot provide the required capability.

| Family | File |
|--------|------|
| MultilineText | `DxUiTests.MultilineText.cpp` |
| Theme | `DxUiTests.Theme.cpp` |
| WindowHost | `DxUiTests.WindowHost.cpp` |
| TextField | `DxUiTests.TextField.cpp` |
| Grid | `DxUiTests.Grid.cpp` |
| ReadOnly | `DxUiTests.ReadOnly.cpp` |
| Animation | `DxUiTests.Animation.cpp` |
| Controls | `DxUiTests.Controls.cpp` |
| Tree | `DxUiTests.Tree.cpp` |
| Rendering | `DxUiTests.Rendering.cpp` |
| ComboBox | `DxUiTests.ComboBox.cpp` |
| Accessibility | `DxUiTests.Accessibility.cpp` |
| Tooltip | `DxUiTests.Tooltip.cpp` |
| NativeTextInput | `DxUiTests.NativeTextInput.cpp` |

## 5. Performance Tests

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

## 6. Viewer Plugin Tests

**Projects:** `Tests\ViewerPETests\` and `Tests\ViewerSqliteTests\`

| Project | Coverage |
|---------|----------|
| ViewerPETests | PE, Web, ImgRaw, Text, Space, VLC viewers — DxUi combo host, long-run stability |
| ViewerSqliteTests | List tables, paged reads, sorting, DxUi host, scrolling, paging, theme, tab traversal |

`ViewerPETests` runs most fresh-process viewer cases with the default
120-second child timeout. Its nested six-cycle shell-combo churn case uses a
dedicated 600-second parent timeout so valid long-run coverage is not killed
before the per-child viewer checks can finish and report their own results.

## 7. File-System Plugin Tests

**Project:** `Tests\FileSystemCurlTests\`  •  **Run:** `.\.build\x64\Debug\FileSystemCurlTests.exe`

Focused deterministic coverage for `Plugins\FileSystemCurl\` helpers that do not require live remote credentials.

| Family | Coverage |
|--------|----------|
| IMAP naming and identity parsing | Preferred `<subject> [uidValidity-uid].eml`, strict operation identity, legacy UID-only diagnostic parsing, malformed-name rejection |
| IMAP subject decoding | RFC2047 Q/B decoding, mixed plain/encoded fragments, UTF-8 emoji, non-UTF code pages, malformed sanitized fragments |
| IMAP mailbox status parsing | `STATUS` fields: messages, recent, uidNext, uidValidity, unseen |
| IMAP capability and stale-identity policy | Exact UIDPLUS parsing plus missing/matching/changed UIDVALIDITY results |
| IMAP safe single-message delete | Fake mailbox proves absent UIDPLUS refusal, UID EXPUNGE success/rejection, target-flag rollback/rollback failure status, and preservation of unrelated deleted mail |
| IMAP properties perf model | Command-count guard proving single-message Properties stays constant with mailbox size |
| IMAP listing metadata repair | Batch plan guard proving large missing summary sets are retried instead of skipped, plus bounded per-listing repair fetch budget coverage |
| IMAP security perf model | STATUS/CAPABILITY validation overhead remains constant with mailbox size |

The executable also supports `--perf` for a lightweight deterministic probe of qualified identity parsing, subject decoding, leaf building, `STATUS` parsing, repair batch planning, and the message Properties/security command-count models.

## 8. Settings Schema And Persistence Tests

**Project:** `Tests\SettingsSchemaTests\`  •  **Run:** `.\.build\x64\Debug\SettingsSchemaTests.exe`

Exercises the shipped Preferences schema and synthetic parser inputs, plus the Common settings persistence
boundary: same-process and real child-process compare-and-swap conflicts, exact revision advancement, malformed
UTF-16 save rejection, uint32 overflow rejection, and save blocking when invalid-file backup cannot preserve the
only recovery artifact.

## 9. RedConfigure Tests

**Project:** `Tests\RedConfigureTests\`  •  **Run:** `.\.build\x64\Debug\RedConfigureTests.exe`

Focused deterministic coverage for RedConfigure page definitions, workspace discovery, theme JSON5 parsing/export/validation, RC parsing/writing, placeholder validation, translation table search/filter/sort, theme preview models, and session export/loading behavior.

## 10. Monitor / ETW Tests

**Project:** `Tests\MonitorTest\`  •  **Run:** `.\.build\x64\Debug\MonitorTest.exe`

Generates 150,000+ ETW trace messages across 3 burst scenarios to validate TraceLogging transport.
Fast targeted guards include `--diagnostics-gate-selftest`, `--scrollbar-model-selftest`, and `--document-model-selftest`.

## 11. Tooling Script Tests

**Run locally/full:** `Invoke-Pester .\Tools\Tests`

**Run in artifact-only CI jobs:** `.\Tools\Run-AllTests.ps1 -Suite CI -SkipBuild`

Suite CI includes the deterministic PluginContractTests, SettingsSchemaTests, and CrashHandlingTests
executables. The pull-request workflow separately performs a Debug ARM64 build-only gate; hosted x64
runners do not execute the ARM64 artifacts.

| File | Coverage |
|------|----------|
| `BuildOutputProcess.Tests.ps1` | Build preflight self-test protection plus creation-time job containment, inherited-handle allowlisting, Unicode/stream fidelity, launch-failure cleanup, direct-child delegation, parallel-runspace and stale-descendant exclusion, owner diagnostics, abandoned-owner contamination, and residual compiler diagnostics |
| `BuildProjectSelection.Tests.ps1` | Project selection and direct vcxproj builds |
| `DocumentationDriftContracts.Tests.ps1` | Source/spec documentation drift guards, including split File Operations popup coverage and archived remediation status |
| `HwndRenderTargetResourcesSourceContracts.Tests.ps1` | Shared HWND render-target/brush lifecycle with FunctionBar and StatusBar policy, invalidation, and teardown retained locally |
| `ModalWindowShellSourceContracts.Tests.ps1` | Shared About/Fatal Error modal owner, message-loop, quit-propagation, and owner-restoration boundaries |
| `MSBuildInvocation.Tests.ps1` | MSBuild invocation planning and diagnostic parsing |
| `MtpLiveCloseout.Tests.ps1` | MTP live closeout wrapper safety, archival, and environment-restoration contracts |
| `PackedFileInfoBufferSourceContracts.Tests.ps1` | Checked packed FileInfo sizing/construction/traversal shared by buffered COM facades while streaming and fixture variants remain local |
| `PluginConfigurationConsolidationSourceContracts.Tests.ps1` | Common plugin-configuration schema/codec reuse and lossless Preferences commits |
| `PluginLifetimeConsolidationSourceContracts.Tests.ps1` | Shared callback generation/drain state, callback-return module-pin transfer, and case-insensitive manager lookup leases |
| `PostedPayloadCoalescingSourceContracts.Tests.ps1` | Queue-head-safe keyed payload coalescing for Compare progress and Find result/progress drains |
| `ProcessStreaming.Tests.ps1` | Process output streaming, logging, and kill-on-close descendant containment |
| `RedSalamanderPluginDeployment.Tests.ps1` | Targeted RedSalamander build repopulates sibling binaries/plugins and plugin language resources; tagged `RequiresBuildToolchain`, bounded, and logged |
| `ResourceLocalizationContracts.Tests.ps1` | Resource placeholder positional-order, satellite placeholder-equivalence, and language-neutral embedded-only string contracts |
| `RunAllTestsPlan.Tests.ps1` | CI/Full runner test-plan enumeration, unified test-sandbox root selection, Pester compatibility, dead-PID stale run cleanup, disk-audit evidence, classifier proof, quarantine planning, and aggregate reporting |
| `SanitizedEnvironment.Tests.ps1` | Child process environment normalization |
| `ShowPerfRuns.Tests.ps1` | Perf-run report parsing and filtering |
| `TestHarnessSourceContracts.Tests.ps1` | Source guards for self-test CLI/error handling, artifact serialization, case listing, input isolation, TestSandbox routing, contained child processes, plugin/viewer contracts, Microsoft Drive credential-bound URL validation/redacted diagnostics, and high-risk regression invariants |
| `TestInventory.Tests.ps1` | Source/run-plan test inventory, project/run-plan set equality, execution-kind coverage, live source-contract classification/replacement queue, FileOperations phase-order integrity, and no checked-in current-total ownership |
| `TestRunArchive.Tests.ps1` | TestRun archive size, machine-profile, curated-evidence, and empty-path policy guards |
| `ThemeDistributionContracts.Tests.ps1` | Theme schema, package-assets, licensing, and distribution contract guards |
| `VcpkgInstallSafety.Tests.ps1` | vcpkg path/triplet safety |
| `VerifyNoProductionGetCursorPos.Tests.ps1` | Production sources avoid `GetCursorPos` outside annotated diagnostics |
| `Versioning.Tests.ps1` | Local build-number reuse/allocation |
| `ViewerChromeSourceContracts.Tests.ps1` | Viewer chrome keyboard routing, launcher subsystem, and Escape contract guards |
| `WingetValidation.Tests.ps1` | Winget validation warning suppression, failure propagation, portable manifest metadata, and VC runtime ZIP helper coverage |

Fast vcpkg merge coverage:

- `.\Tests\vcpkg-merge-synthetic-test.ps1` — synthetic lock/merge coverage.

Manual-only intrusive validation:

- `.\Tests\vcpkg-merge-lock-validation.ps1` — intrusive vcpkg install/lock validation
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

`SettingsSchemaTests` also owns Settings Store connection-identity recovery proof: strict reload rejects case-colliding
IDs, startup recovery assigns distinct canonical replacements without copying ambiguous saved-secret references, the
repaired snapshot persists through source CAS, and subsequent strict reload succeeds. The focused Commands case
`connection_secret_authorization_scopes` covers secret-kind/purpose isolation, timeout zero, expiry, unsigned tick-wrap
arithmetic, background continuation, targeted revocation, and the same global session-clear path used for lock/logoff.

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
