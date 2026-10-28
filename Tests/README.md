# Running only the tests your change affects

File-operation presentation fixtures declare their producer capabilities explicitly. Copy tests
cover written bytes without logical child counts as well as supported copied-file/folder counts;
generic item-overflow checks use an operation that exposes item scopes. Monitor source-contract
checks use the exact clean external pin through the shared resolver, never a removed local source.

```powershell
./Test-Changes.ps1 -Explain
./Test-Changes.ps1 -Configuration Debug
./Test-Changes.ps1 -Mode PrePush -Configuration Debug
./Test-Changes.ps1 -Mode Full -Configuration Debug
```

The default uses local merge-base changes plus staged, unstaged and untracked work. Read the explanation before a large run. Unknown code/build inputs widen coverage. Selected work is partial evidence. PR delegation is pending CI work, and only an enabled workflow with matching reviewed candidate bytes/profile can justify it.

Native test files use `Scope.Tests.Something.h/.cpp`; register new active files in `native-test-files.json`. `test-scopes.json` owns standalone scope/PR metadata. Shared-library or runtime byte changes conservatively invalidate local standalone reuse. No consumer dependency pin changes are part of this workflow.

For independent resource investigations, force fresh execution and retain paired fixture/build provenance. Timing, native ARM64, real devices, IME/assistive technology and foreground input remain separate qualification requirements. These commands do not manufacture those claims.
```powershell
./Test-Changes.ps1 -Suite Commands -CommandsFamily file-operations -SkipBuild
./Test-Changes.ps1 -Mode Full -Interactive -ResumeFrom <run-id>
```

Execution and evidence remain owned by Tools/Run-AllTests.ps1. Exit 2 means incomplete repository coverage even when selected assertions pass. Focus-taking work is pending unless Interactive is requested after agreement to the time; a selection made only of pending work executes nothing, lists it as PENDING_INTERACTIVE and exits 2. Test paths remain beneath the existing marked X:\RedSalamander.Perf root. The wrapper does not authorize root initialization. Tools/Get-TestInventory.ps1 is the live coverage inventory.

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
| In-product self-tests | `SelfTest` | Runner-native case listing plus `Tools\Modules\Testing\TestRunPlan.psm1` |
| Native component/plugin tests | `Executable` | `Tests\*.vcxproj` reconciled with CI/Full run plans |
| Performance tests | `CppUnitTest` | `PerformanceTests2.vcxproj` plus CI/Full run plans |
| Tooling tests | `Pester` / `PowerShellScript` | CI/Full run plans plus source-derived manifest |

Related specifications:
- `Specs/Testing/Testing_SelfTests.md` — result contract
- `Specs/Testing/Testing_TestCoverage.md` — coverage policy and risk map
- `Specs/Testing/Testing_PerformanceValidation.md` — perf validation

## Quick Start

```powershell
.\Tools\Run-AllTests.ps1                      # Build Debug + run ALL self-tests
.\Tools\Run-AllTests.ps1 -Suite PR             # Build Debug + run the GitHub Actions pull-request/push gate locally (about ten minutes of tests)
.\Tools\Run-AllTests.ps1 -Suite CI             # Build Debug + run the nightly suite locally (adds the in-product self-test suites)
.\Tools\Run-AllTests.ps1 -Suite Full           # Build Debug + run self-tests, native tests, PerformanceTests2, and local script tests
.\Tools\Run-AllTests.ps1 -Suite Commands       # Single suite
.\Tools\Run-AllTests.ps1 -Suite Commands -CommandsFamily file-operations # One governed family process; not Full evidence
.\Tools\Run-AllTests.ps1 -SkipBuild -FailFast  # Fast iteration
.\Tools\Run-AllTests.ps1 -TimeoutMultiplier 3  # Slow machines
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Resume -ResumeFrom <run-id> # Exact promoted reuse only
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Affected -ImpactBase origin/HEAD # NOT_EVALUATED iteration
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
**Source:** `RedSalamander\SelfTest\Commands\Commands.Tests.Runner.cpp` + 13 `.cpp` family files
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
| BatchRename | `SelfTest\Commands\Commands.Tests.BatchRename.cpp` | Batch Rename dialog source guard |
| PluginConfig | `SelfTest\Commands\Commands.SelfTest.PluginConfig.cpp` | Plugin configuration, file-system plugin |
| Connections | `SelfTest\Commands\Commands.SelfTest.Connections.cpp` | Connection manager, credential prompts, plugin-backed MTP picker, session-first credential persistence failure |
| Preferences Dispatch | `SelfTest\Commands\Commands.Tests.Preferences.Dispatch.cpp` | Preferences dialog — all categories, DxUi surfaces |
| CompareOptions | `SelfTest\Commands\Commands.SelfTest.CompareOptions.cpp` | Compare directories options, progress |
| Search | `SelfTest\Commands\Commands.Tests.Search.cpp` | Find dialog, local search index, quick search/filter |
| Shortcuts | `SelfTest\Commands\Commands.Tests.Shortcuts.cpp` | Shortcuts window grouped runner |
| ViewCommands | `SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp` | View commands, selection, sort, pane, tabs, FolderView rendering-alert persistence, draw-item brush reuse, refresh-to-paint telemetry, IconCache live-failure retry, DPI repaint |
| FileOps | `SelfTest\Commands\Commands.Tests.FileOps.cpp` | File operations issues pane, speed limit |
| Navigation | `SelfTest\Commands\Commands.Tests.Navigation.cpp` | Navigation location, GoTo |
| Dialogs | `SelfTest\Commands\Commands.SelfTest.Dialogs.cpp` | About, fatal error, splash, change case, rename, filter, mask |
| ShellCommands | `SelfTest\Commands\Commands.Tests.ShellCommands.cpp` | Shell-integrated pane commands, bounded FolderView drop parsing/targeting/effect semantics, pointer target integrity, and verified move-clipboard invalidation |

## 2. Compare Directories Self-Test Suite

**Flag:** `--compare-selftest`
**Source:** `RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.Tests.Runner.cpp` + 4 included case files
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
**Source:** `RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.Tests.Runner.cpp` + 4 included phase files
**Inventory:** `RedSalamander.exe --selftest-list-cases --fileops-selftest`
emits setup, the active ordered phases, and cleanup. The source-derived manifest
validates the phase order against the `Step` enum without freezing a total here.

Async tick-driven state machine testing file copy/move/delete operations end-to-end.
Organised into 12 families spanning phases 5–16.

| Family | Coverage |
|--------|----------|
| Phase 05 — Discovery | Streaming discovery (single traversal feeds execution and closes totals), Skip discovery, cancel latency and capacity release, queued/queue-paused cancel and cleanup drain, Queue/Parallel mode switching during discovery |
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
| Notes (I35 S1) | Questions, user skips and expected outcomes are Information notes, not warnings; a real failure still counts; the informational-notes setting changes only display; inherited permissions, protected DACLs and lost streams classify alike on the native and bridge Copy routes; an Overwrite keeps named streams and MOTW |

---

Nested archive Copy regressions are in `FileOpsFamily_CrossProviderRoutes`:
`FileOps_ArchiveNestedCopyLocal` and `FileOps_ArchiveNestedCopyUnc` copy entries from ZIP,
7z and multipart RAR through the real File Operations bridge. They verify exact bytes and
one destination leaf for canonical, Windows-joined and backslash source paths, while
invalid destination names remain rejected. The UNC case uses the existing sandbox-only
loopback alias and explicitly skips when the administrative share is unavailable.
After agreeing a time for the runner's desktop lease, run the two cases with
`./Test-Changes.ps1 -Suite FileOps -CaseFilter FileOps_ArchiveNestedCopy -Interactive`.
The ordinary PluginContractTests capability pass also checks archive leaf/parent/relative
path derivation and rejection of separators in proposed child names without desktop input.
Per-copy `FileOps.SelfTest.ArchiveNestedCopy` timings are correctness evidence, not a
large-file or NAS throughput benchmark.
The small RAR fixture's pinned upstream provenance, hashes and license are in
`Plugins/FileSystem7z/Tests/MultipartRar/README.md`. Only its two small regular-file
entries are read; the large compression fixture and link entry are never copied.

## 4. Product UI and shared DxUi tests

`Tests/ProductUiTests/` contains product helper, palette, viewer adapter and lifetime regressions.
The canonical Full runner builds and executes it against the exact external DxUi pin. RedConfigure
and in-product Commands tests qualify actual windows, localization, focus, accessibility and teardown.

Shared controls, native/embedded hosting and gallery tests belong in the
[DxUi repository](https://github.com/RedSalamanders/DxUi). Run its `test.ps1` in the required profiles;
use its `gallery.ps1` for shared control images. The old in-tree DxUiTests project is retired.
Neither library-only results nor a ProductUiTests pass replaces Full product qualification.

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

ViewerImgRaw fresh-open probes wait for the UI-applied final result, independently of
host alert activity. The TIFF orientation fixture gates its real provider read to prove
that startup alert clearing cannot satisfy completion, then checks the same decoded
image, orientation and geometry after release. PNG and resource-policy probes share
that terminal predicate; their eight-second decode deadlines remain unchanged.

| Project | Coverage |
|---------|----------|
| ViewerPETests | PE, Web, ImgRaw, Text, Space, VLC viewers — DxUi combo host, long-run stability |
| ViewerSqliteTests | List tables, paged reads, sorting, DxUi host, scrolling, paging, theme, tab traversal |

Debug `TestViewerVlcLongPathsAndErrorMessages` uses silent sandbox WAV files at
259, 260, and 275 characters and proves the decoded duration through the installed
VLC backend. It covers F3/embedded mode, Unicode/URI delimiters, extended and
forward-slash paths, missing files, directories, non-local paths, missing installs,
and a found install with an invalid DLL. Run it without taking desktop focus:
`ViewerPETests.exe --no-activate TestViewerVlcLongPathsAndErrorMessages`.
Actual decoding is explicitly skipped when a standard VLC installation is absent;
error-category assertions still run. The test never changes Windows long-path policy.
Guarded links within the same sandbox also cover long-path targets and broken
targets; only missing OS symbolic-link privilege skips these conditional cases.

`TestViewerSpaceAnalysisHostVirtualGridAndLifetime` checks the real grid/filter,
resize/collapse and detached-provider path behind the no-activation guard.
Its workspace assertions cover exact discovered/original/shown exclusion accounting, restore,
separate cleanup marks/review/recorded host handoff, configuration defaults/roundtrip,
and Unicode/unknown-fact JSON plus CSV formula-prefix protection. The recording host never deletes data. The same case drives the production export preparation/writer through a test-only synchronous gateway that admits only marked-sandbox paths. It checks exact JSON rows/totals, CSV metadata/UTF-8 identification, failed atomic publication preserving an existing file, and bounded Close while an owned duplicated-event gate holds the module-pinned writer. The gateway, fault and gate are absent from test-disabled builds.
The focused interactive case exercises the Layout button/menu, Ctrl+1/Ctrl+2 and Escape
under the existing warning/desktop lease, after the user agrees to the time.
It verifies visible file/folder Modified timestamps, unavailable missing/invalid provider dates,
and the complete timestamp through narrow/wide/narrow layouts. The source-callback baseline is
sampled immediately before Detach, after the desktop witness drains pending live-host events;
Detach and stale UIA access must still leave that counter unchanged.
`TestViewerSpaceAnalysisAllowsAmbientDesktopChanges` runs the same UIA scenario with
read-only simulated cursor/foreign-foreground observations; it never generates input.
`TestViewerSpaceDesktopObservationContracts` covers fail-closed guard/capture states,
transient owned foreground, nonmonotonic/unavailable last-input ticks, nested scope
refusal, repeated event-drain barriers and restart/teardown. For an exact contained
nonactivating run use `ViewerPETests.exe --group=noninteractive <case-name>` with the
marked test sandbox configured. The desktop witness records sample/change counts
and drain time; it does not require the person to stop using the mouse, and never
restores their foreground or pointer. A focused pass does not qualify the whole group.
ViewerSpace lifecycle captures use the same witness and label their provenance
`foreground_witness_scope=guarded-process-owned-events`; the compatibility
`foreground_preserved` flag records absence of owned activation through the frame
barrier, rather than unchanged desktop-global identity.
The gated refresh capture waits for `Scanning` plus an empty pending-update queue;
negative controls reject still-queued snapshots without weakening capture stability.

`ViewerPETests` runs most fresh-process viewer cases with the default
120-second child timeout. Its nested six-cycle shell-combo churn case uses a
dedicated 600-second parent timeout so valid long-run coverage is not killed
before the per-child viewer checks can finish and report their own results.
Both viewer harnesses expose `--group=noninteractive` and
`--group=interactive`. The noninteractive groups install the shared test-only
activation blocker; only combo/keyboard/prompt ViewerPE cases and ViewerSqlite
tab traversal run in the interactive group.
ViewerPE also accepts an exact registered name after its group, for example
`ViewerPETests.exe --group=interactive TestViewerSpaceWindowOpensWithoutVisibleChildFallbackAndEscapeCloses`.
This runs only that case through the fresh-process parent, including containment
and foreground restoration. Run it inside the existing runner desktop lease
after agreeing to the time. A missing name or activation-class mismatch fails
without launching a child; focused coverage does not qualify the whole group.

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

**Run locally/current:** `Invoke-Pester .\Tools\TerminalEngine\CurrentToolingProfile.Tests.ps1`

This is the same dynamically discovered current profile used by CI and Full.
It deliberately excludes the SHA-256-sealed Gate0/Round4 Terminal cohort; see
`Tools/TerminalEngine/README.md` for explicit archival reproduction commands.

**Run in artifact-only CI jobs:** `.\Tools\Run-AllTests.ps1 -Suite PR -ValidationMode Fresh -SkipBuild`
(pull requests and pushes) or `-Suite CI` (nightly).

Suite PR and Suite CI both include the deterministic PluginContractTests, SettingsSchemaTests, and
CrashHandlingTests executables. Pull requests build and test x64 Debug and ARM64 Debug on native
runners (`windows-2025-vs2026`, `windows-11-vs2026-arm`); the ARM64 artifacts run on the ARM64 host.

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
| `SpecInformationArchitecture.Tests.ps1` | Specification authority classification, exact WIP indexing, protected Done history, Markdown link/anchor parsing, normative-plan leakage, consistency-ledger completeness, and retired tombstone contracts |
| `TestHarnessSourceContracts.Tests.ps1` | Source guards for self-test CLI/error handling, artifact serialization, case listing, input isolation, TestSandbox routing, contained child processes, plugin/viewer contracts, Microsoft Drive credential-bound URL validation/redacted diagnostics, and high-risk regression invariants |
| `TestInventory.Tests.ps1` | Source/run-plan test inventory, project/run-plan set equality, execution-kind coverage, live source-contract classification/replacement queue, FileOperations phase-order integrity, and no checked-in current-total ownership |
| `TestRunArchive.Tests.ps1` | TestRun archive size, machine-profile, curated-evidence, and empty-path policy guards |
| `ThemeDistributionContracts.Tests.ps1` | Theme schema, package-assets, licensing, and distribution contract guards |
| `VcpkgInstallSafety.Tests.ps1` | vcpkg path/triplet safety, offline dependency-pin contract, MSBuild install overlay pin |
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
| `--selftest-no-activate` | Explicitly require the no-activation guard for Compare Directories or File Operations; these focused suite launches also default to the guard |
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

The shared support and sandbox-path headers accept compiler-provided Windows
feature macros without redefining them. `RelocatedTestSandbox.Tests.ps1` builds
and executes this header-only path outside the checkout with warnings as errors;
its ASan Debug fixture also requires actual sanitizer instrumentation.

`SettingsSchemaTests` also owns Settings Store connection-identity recovery proof: strict reload rejects case-colliding
IDs, startup recovery assigns distinct canonical replacements without copying ambiguous saved-secret references, the
repaired snapshot persists through source CAS, and subsequent strict reload succeeds. The focused Commands case
`connection_secret_authorization_scopes` covers secret-kind/purpose isolation, timeout zero, expiry, unsigned tick-wrap
arithmetic, background continuation, targeted revocation, and the same global session-clear path used for lock/logoff.

Captured child execution uses `TestSupport/TestSupport.Tests.ChildProcess.h`. Callers pass an
executable plus structured arguments; the runner owns all process, thread, pipe,
and JobObject handles through WIL, starts the root suspended, allowlists inherited
standard handles, assigns the kill-on-close job before resume, and drains stdout
and stderr concurrently even after their retained prefixes reach their configured
limits. Timeout and cancellation contain the whole process tree. A nonzero exit is
a completed launch whose exit code and separate streams remain available to the
test oracle.

Mounted-volume coverage uses only the test-owned VHDX that an elevated
`Tools\Run-AllTests.ps1` creates, proves and mounts inside the current run. Native
tests never mount a volume; `TestSupport/TestSupport.Tests.OwnedVirtualDisk.h`
validates the handed-off mount and refuses any other volume. Unelevated runs skip
the mounted routes with the harness reason.

Tests never mount, link (junction, symbolic link, hard link), share, `subst` or
otherwise expose a directory or volume outside a marked sandbox. Every link a test
creates lives in, and targets a path inside, the same marked sandbox: native code
calls `Common::Testing::RequireSandboxedLinkTarget` first (or uses the guarded
`CreateSandboxedSymbolicLink` / `CreateSandboxedHardLink`), SUBST and loopback-share
aliases call `RequireSandboxedExposureTarget`, and PowerShell uses
`New-RSTestSandboxLink`. A refusal fails the case. Sandbox cleanup never follows
links or mount points. `Tools\Tests\TestHarnessSourceContracts.Tests.ps1` rejects
any such call in test code outside its documented allowlist.

### Artifacts

When launched through `Tools\Run-AllTests.ps1`, artifacts are written under
`X:\RedSalamander.Perf\runs\<runId>\artifacts\selftest\last_run\`. `X:` defaults
to the repository drive and may be any fixed local drive selected with `-TestRoot` or
`REDSALAMANDER_TEST_ROOT`; the directory name is exact. First initialization requires
`-AllowExternalTestRoot`, later use requires the ownership marker, and arbitrary roots,
descendants, UNC paths, `.build\TestSandbox`, and historical `RSPerf` names are rejected.
The runner redirects `TEMP`/`TMP` into the current run and restores caller environment
values. After children and owned-disk teardown finish, it verifies and renames the complete
run on the same volume into `evidence\completed-runs\<runId>`, retaining passed, failed,
skipped and early-exit output. Summaries and dashboards point to that final location;
native trace/result bytes retain their original paths, mapped by create-only records in
`evidence\run-archive-records\<runId>`. Startup archives abandoned parseable runs instead
of deleting them; live owners, unsafe links, unresolved disks and conflicting destinations
remain protected. Ordinary runs inspect only the selected root; explicitly enabled
alternate-volume runs also finalize their same-id participating marked roots.
Normal runs never perform legacy cleanup outside the selected root;
`Tools\Clean-TestSandbox.ps1` remains a separate dry-run-by-default migration tool.
Disk audits require a valid marked root even when a requested path is absent; absence
does not establish cleanliness. To maintain a test machine, first inventory the exact
reported paths and preserve useful results, failures and diagnostics under the same
root's `evidence`. Verify preserved bytes and unchanged source identities before removal.
Then preview one exact run with `Tools\Clean-TestSandbox.ps1 -TestRoot X:\RedSalamander.Perf -RunId <id>`;
only an explicitly reviewed `-Apply` removes it. Apply protects live runner/harness owners
with the shared PID-reuse checks; failure returns `Status: Failed` and a nonzero exit.
Restore selected archived files into fresh contained scratch, preserving the original
receipt; never extract over the active root or automatically recreate recorded links.
Manually named runs need explicit review and are never automatically swept by naming
prefix. Empty misplaced root remnants require containment checks and nonrecursive removal.
See [the normative maintenance contract](../Specs/Testing/Testing_SelfTests.md).
Tooling Pester tests that create scratch files or synthetic run trees use
`Tools\Modules\Testing\TestRunPlan.psm1` `New-RSTestSandboxScratchDirectory(...)` and write under
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
| [Testing_TestCoverage.md](../Specs/Testing/Testing_TestCoverage.md) | Coverage policy, required risk domains, and inventory authority |
| [Testing_PerformanceValidation.md](../Specs/Testing/Testing_PerformanceValidation.md) | Performance validation requirements |
| [Testing_SelfTestRemoteCredentials.md](../Specs/Testing/Testing_SelfTestRemoteCredentials.md) | Remote storage credential setup |


## File Operations documentation gallery

The public examples come from the private English-language fixture in
`FileOps_ReviewRemediationContracts`. Its `fileops/documentation` artifacts cover
ten scenes: running progress, open counting, totals-only Move, file-only completion,
paused/queued/completed Copy, destination conflict, the complete wrapped terminal
note and a selected Issues row. It seeds invented Workshop media names/paths, opens
none of those paths, and asserts exact count wording and native count/note layout
before capture:

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -CaseFilter FileOps_ReviewRemediationContracts -NonInteractive -Configuration Debug -Platform x64 -TestRoot D:\RedSalamander.Perf
```

Use the broader presentation gallery below for additional states. The shared
[screenshot contract](../Specs/Testing/Testing_SelfTests.md#documentation-screenshot-scenarios)
applies to every published frame, including titles and breadcrumbs.

Run the real application renderer with named synthetic scenarios, without taking desktop focus:

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -CaseFilter FileOps_VisualGallery -Platform x64 -Configuration Debug -TestRoot C:\RedSalamander.Perf
```

Use an already marked test root; the usual explicit first-initialization rule applies otherwise.
The runner prints the run ID. PNGs and `scenarios.tsv` are in
`<test-root>/evidence/completed-runs/<run-id>/artifacts/selftest/last_run/fileops/visual-gallery/`
after the runner exits (under `runs` while active).
The case is intentionally absent from the broad FileOps suite; select its exact name.

Add presentation scenarios in
`RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.Tests.VisualGallery.cpp`.
Use coherent snapshot facts and the existing engine policy. The test-only fixture affects no
ordinary popup when it is unset. Extend representative theme/width coverage in that same case.
`Tests/TestSupport/TestSupport.Tests.WindowScreenshot.h` supplies complete owned-window capture for other UI fixtures.
It needs a Windows session with Windows Graphics Capture and D3D11 available; it never captures the desktop.
Layered fixtures can supply `WindowScreenshotColorProbe` witnesses in client pixels, with the expected
theme color, tolerance and minimum pixel population. The throughput fixture requires both text and
background in its legacy label region. On rejection the helper saves a diagnostic PNG and requests
a non-activating repaint with capture active, for at most three frame attempts. Call with probes
from the window's UI thread. Persistent mismatch returns `ERROR_INVALID_DATA`; encoder success
alone is insufficient. Probes establish the selected region's presence, not full layout correctness.

Inspect every output before publishing documentation. Copy accepted images with a manifest that records
source revision, dirty source identity, executable identity, scenario/theme/DPI/width and image hashes.
Keep generated proposal mockups separate. This gallery is presentation evidence; retain normal engine,
interaction and accessibility tests for behavioural claims.

The gallery now uses French UI resources, long French text, both available monitor DPIs, real
confirmation overlays, Issues views, destination history and recovery choices. For explicitly
authorized hover/focus/pressed and same-window DPI transition captures, select the separate lane:

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -CaseFilter FileOps_VisualGalleryInteraction -Configuration Debug -Platform x64 -TestRoot C:\RedSalamander.Perf
```

This exact case takes the runner desktop lease and displays the shared warning before briefly
using the foreground. It restores cursor/foreground state. Static capture remains non-activating.

The supplied-review safety cases use `ProductUiTests` (directory decoding and destination-bound payloads),
`-Suite FileOps -CaseFilter FileOps_ReviewRemediationContracts -NonInteractive` (parent pinning and typed prerequisite results),
and exact noninteractive `FileOps_SmallFileCopyCost` / `FileOps_CrossRootLocalCopy` selections.
XP-09 reports an unreachable share as skipped, with no cross-root qualification. The small-file
benchmark checks actual names and bytes on both routes outside its timed interval and exercises
missing/corrupt output controls. Broad FileOps selections retain the desktop lease.

The exact no-activation `FileOps_SelectionShapeLocal` and `FileOps_SelectionShapeSmb` cases compare
folder versus equivalent selected-file Copy/Move, actual file overlap, fresh count convergence,
live popup totals and exact contents. Local controls include Auto with inactive Manual one, Manual
one/four and a 4,097-file selection. The SMB case uses the configured `path.smb.primary` resource,
native/bridge Copy, managed Move, a bounded nested thirteen-file/six-directory fixture and source-cleanup cancel/Retry/Skip/unknown
controls. An unavailable SMB resource is a skip, never real SMB qualification. Use separate exact
case filters with `Tools/Run-AllTests.ps1 -Suite FileOps -NonInteractive`; broad families keep their
interactive lease. See `Specs/Testing/Testing_PerformanceValidation.md` for metrics and acceptance.

## Product documentation captures

Run the exact `cmd_product_documentation_gallery` Commands case with
`Run-AllTests.ps1 -Suite Commands -CaseFilter cmd_product_documentation_gallery -NonInteractive`.
It captures 26 private workspace, Search, Preferences, plugin, connection and palette
scenes without global input. `ViewerPETests TestViewerDocumentationGallery --no-activate`
captures eight generated-content viewer scenes; `TestViewerSpaceLifecycleCaptures --no-activate`
covers eleven Space states. Preserve the build/runtime receipts, scenario manifests
and complete frames, then perform full-frame privacy review before publication.
These are focused presentation checks; ordinary interaction/profile obligations remain.

The product gallery now captures 45 scenes (26 original plus 19 supplemental).
Its modal utilities always cancel; Terminal uses one-shot shell-free VT input;
provider browsing uses owned loopback APIs or RAM MTP data.
`ViewerPETests TestViewerVlcDocumentationPlayback --no-activate` adds two real
decoder playback/pause captures from a generated silent AVI. An installed VLC
runtime is required for publication; a missing-runtime skip cannot close this gate.

F7 checks use process-owned foreground events and a bounded event drain, with foreign activity reported as diagnostic counts. The immutable local DxUi selection-fix pin is under qualification; its original parent and unpublished status are recorded in the active toolchain plan. Publishing and clean remote restore remain separate merge/release gates.

Compiler-policy remediation preserves failed receipts and strict archival checks. Clipboard fixtures use independent synthetic sequence identities per invocation and retain a duplicate-admission negative control. Compare link fixtures clean only their owned link entries and check exact target contents. Find selection fixtures establish the Grid's current row through public selection requests before keyboard context-menu tests. Preferences preview coverage injects a delayed theme notification after the next draft edit; footer/pointer failures retain actual focus, geometry and animation predicates.

The compiler qualification regressions preserve completion predicates rather than assuming one bounded pump drains the queue: first Compare body paint, posted thumbnail size command, settled task-card height and warm FolderView quiescence. VLC's unchanged 900 ms publication bound excludes trace/focus-probe administration and records it separately; the clock begins before the actual focus command. The unknown-byte paused popup scenario must stop indeterminate animation without pretending that an unknown total is determinate.

Paused unknown-progress cadence uses a separate FileOps run with CaseFilter FileOps_PopupIdleCadence (or the single FileOpsFamily_PopupIdleCadence family). Unlike Commands, the native FileOps selector does not accept a comma-separated cross-suite case list. It checks that the actual hosted bar remains enabled and indeterminate with animation paused, alongside the unchanged waiting repaint bound. ViewerVLC switch timing retains its 900 ms deadline and records individual owner-UI playback API stages to attribute failures.

The configured VLC preview fixture holds media preparation off the owner thread, proves both original 900 ms publication predicates while held, coalesces the latest selection without another worker, and verifies actual instance/player identities after release. Its decoder-ready wait is separate from UI responsiveness.

TestViewerVlcLongPathsAndErrorMessages also exercises real-player media-submit refusal, failed completion delivery, failed-post coalesced recovery, keyboard Stop, and Close with the worker held. It verifies 100 ms UI returns, no stranded ownership or stale playback, retained hidden HWND ancestry, and actual failure counters before recovery.

The held real-player media controls also force the delayed loading overlay and assert the actual HUD WS_VISIBLE style, nonempty layout and top sibling position before invoking Stop. This catches a logical Stop-available flag whose real HUD is hidden or covered.

The synthetic-root search_service_sqlite_startup_warms_overridden_roots lifecycle fixture uses the existing case-scoped traversal-seed control, restores it, retains its 15-second deadline, and explicitly checks currentness-unproven persisted states before the exact-result query. Native MFT/journal coverage remains in its dedicated default-path cases. A global forced-traversal diagnostic is never Full qualification.

The next compiler-remediation candidate aligns FileOps fixtures with typed disk-full actions, Cancel defaults, primary Copy label placement and per-object deletion accounting. Commands junction cases must clean their own links and preserve target bytes. Gallery foreground and ViewerSpace inspector failures now include bounded diagnostics; their causes and final qualification remain open until actual focused and Fresh Full receipts pass. The retained P3 Debug Full failed and its raw run could not be archived because three Commands junctions remained. Passing child messages do not override the failed ViewerPE interactive group.

The focused P4 junction group now passes all six observations and archival; Copy label placement and exact deletion accounting pass too. The 7z provider fixture follows the existing `/` plus `\` path contract, with preferred slash and case sensitivity retained. Gallery diagnostics cover every capture and final event drain, while splitter failures record logical/physical pointer and hit-window data and consent failures record the actual typed facts. The independent-volume case requires an existing marked alternate root declared in machine resources and explicit `-AllowAlternateVolumeTestRoot`; a blocked environment is not a passing observation or a reason to weaken its assertions.

The consent matrix's source-retention branch runs as Move, with an independent EFS Copy action-refusal control. Fake-MTP discovery owns a private profile through asynchronous work and restores the caller exactly. New parent-message and artifact-rotation regressions first run against the old algorithms; expected red receipts remain separate from final qualification. The rotation fixture is `cmd_selftest_artifact_rotation_readonly_link_and_locked_failure`.
