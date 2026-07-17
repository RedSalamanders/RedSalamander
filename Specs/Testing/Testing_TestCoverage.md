# Test Coverage Specification

## Overview

This document describes coverage domains and durable test contracts. It does not
own mutable current totals.

Use `RedSalamander.exe --selftest-list-cases` for the authoritative in-product
case inventory. Use `Tools\Get-TestInventory.ps1 -Format Json` for the
source-derived manifest of static registrations, native test projects, CI/Full
run-plan entries, and execution kinds. The manifest derives project-backed
surfaces from `Tests\*.vcxproj` plus `Tools\TestRunPlan.ps1` and fails when the
sets or kinds diverge. It also parses every `TestHarnessSourceContracts` `It`
block and classifies it as `LexicalSafetyBan`, `StructuralGraphInvariant`,
`BehavioralShadow`, or `MixedSourceShape`; the latter two form the live runtime-
replacement queue. Current totals must not be copied into this specification;
archived run counts below remain dated evidence rather than inventory authority.

The Compare delete-burst maintenance parity case launches its isolated foreground
service without consuming the queued-state observation in a generic readiness
probe. Its predicate owns the first observation, uses the shared 30-second cold
service readiness budget for each bounded maintenance window, and reports
captured child-process exit output when the service terminates early. This avoids
misclassifying an in-progress automatic compaction window as a missing-pipe
failure under broad suite load.

Recent focused coverage updates:

- 2026-07-17 Observatory Track 17 inventory checkpoint: `Tools\TestInventory.ps1` now parses and uniquely
  classifies every live `TestHarnessSourceContracts` case, exports per-case line/assertion metadata and category
  totals through `Get-TestInventory.ps1 -Format Json`, and exposes behavioral/mixed cases as the replacement queue.
  The classifier found 15 structural graph invariants and 133 behavioral/mixed source-shape checks in the current
  148-case file; these are an observed checkpoint, not frozen specification totals. `TestInventory.Tests.ps1`
  passes **7/7**, including classification coverage, uniqueness, category accounting, JSON parity, native project
  parity, run-plan coverage, and FileOperations phase-order integrity.

- 2026-07-17 Observatory Track 14 lifecycle/UI closeout: the settings Commands case now proves
  `BeginProcessShutdown()` admits one final snapshot, rejects a crossing normal submitter, and makes duplicate
  finalization idempotent; the NavigationView Commands case posts deliberately stale suggestion results before
  Escape, apply, and exit; and the DxUi Menu suite exercises cached geometry and viewport-only paint with 4,096
  rows. Debug `RedSalamander` and `DxUiTests` builds completed with zero warnings/errors, both exact Commands cases
  exited 0, and the focused Menu suite passed. Archived UI/perf evidence is under
  `Specs/TestRuns/4cb089111a23/UI/2026-07-17_132300_observatory_track14_lifecycle_navigation_menu/`.

- 2026-07-17 Observatory Track 9 non-provider reentrancy checkpoint: `MenuBar` callback root replacement,
  `NativeMenuBarHost` destruction from inside a nested DxUi popup loop, Tree selection-delegate root replacement and
  stable-ID expander re-resolution, and NavigationView full-path owned-window activation now have focused regressions.
  Debug `DxUiTests` built with 0 warnings / 0 errors; Control, Tree, and Menu suites exited 0; the serialized Debug
  `RedSalamander` project build exited 0; and
  `cmd_pane_navigationView_full_path_popup_owned_window_activation` exited 0. Source contracts passed **142 / 0 / 0**.
  The pre-existing `cmd_pane_navigationView_full_path_popup_edit_route` Escape focus-return failure reproduced with
  the owned-window block skipped and is preserved separately under
  `Specs/TestRuns/SINON/Continuation/2026-07-17_0938_observatory_track9_full_path_focus_failure/`; it is not counted
  as proof for or against the isolated owned-window change.

- 2026-07-17 Observatory Track 9 text-input checkpoint: focused replacement now synchronizes native/TSF and
  accessibility state before terminal notification, all text-change callers stop if the callback destroys the
  control, window/app deactivation restores active IME preview base state, and masked caret/range/selection/hit-test
  geometry maps source UTF-16 boundaries through a cached user-visible text-element map. Debug `TextField`,
  `NativeTextInput`, and `Accessibility` suites exited 0; source contracts remained **142 / 0 / 0**. The archived
  TextField perf capture at
  `Specs/TestRuns/SINON/Continuation/2026-07-17_observatory_track9_text_input/text_field_perf.jsonl` contains six
  `dxui.textinput.masked_index_map_rebuild_us` samples (96 us minimum, 231 us median, 1,140 us maximum); the seven-
  UTF-16-unit/three-text-element ZWJ fixture rebuilt in 155 us and reused the retained map for later geometry.

- 2026-07-17 Observatory Track 9 UIA closeout: retained Tree-item providers now keep the full 64-bit stable item
  ID across reorder, preserve their runtime ID, re-resolve sibling/action targets, and return
  `UIA_E_ELEMENTNOTAVAILABLE` after removal. Text ranges normalize Character, Word, logical Line, wrapped visual
  Line, and Document units; Character expansion keeps a ZWJ emoji intact, and wrapped Line expansion is covered
  through the cross-thread host dispatch path. The serialized Debug `DxUiTests` project build completed with 0
  warnings / 0 errors, the focused `Accessibility` suite exited 0, and source contracts passed **142 / 0 / 0**.

- 2026-07-16 Observatory IMAP identity/delete-safety checkpoint: FileSystemCurl now lists messages as
  `<subject> [uidValidity-uid].eml`, revalidates UIDVALIDITY before fetch/delete, requires exact UIDPLUS capability
  before marking a message deleted, never issues mailbox-wide EXPUNGE, and compensates a rejected UID EXPUNGE by
  removing the target's `\Deleted` flag while preserving the original error and cleanup status. The deterministic
  fake mailbox covers absent UIDPLUS, unrelated pre-deleted mail, UID EXPUNGE success/rejection, rollback success/
  failure, and stale UIDVALIDITY. Debug test/plugin builds and the test-enabled Release test build had 0 warnings /
  0 errors; FileSystemCurlTests passed in Debug and Release; source contracts passed **131 / 0 / 0**. Release CPU
  and constant command-count evidence is archived at
  `Specs/TestRuns/4cb089111a23/FileSystemCurl/2026-07-16_2028_observatory_track2_imap_safety/`; live network latency is
  explicitly unclaimed pending a deterministic network-capable IMAP fixture.

- 2026-07-16 Observatory release fail-closed checkpoint: `Tools/ReleaseArtifactPolicy.ps1` and
  `Tools/Tests/ReleaseWorkflowPolicy.Tests.ps1` now prove the exact x64-only and x64+ARM64 portable/MSIX matrices,
  reject missing requested legs, wrong PE architecture, wrong MSIX identity/version/architecture, unexpected files,
  and checksum tampering, and source-check the release workflow's dependency semantics, deterministic upload paths,
  immutable third-party action pins, removed mismatched Squad workflows, Dependabot entry, and CODEOWNERS guard.
  The focused suite passed **13 / 0 / 0**; the complete `Tools/Tests` Pester surface passed **309 / 0 / 0** in
  3m00s. Workflow-only remediation required no native rebuild or repeated CI/Full run.

- 2026-07-15 Lighthouse shared-test-support checkpoint: first-party standalone
  harnesses now share `Tests/TestSupport/TestSupport.h` for exact environment
  restoration, TestSandbox run/scratch/artifact routing, bounded message pumping,
  and typed snapshot polling. Compare's captured child execution now delegates to
  `Tests/TestSupport/ChildProcess.h`, which allowlists inherited handles, assigns a
  suspended root to a kill-on-close JobObject before resume, concurrently drains
  bounded stdout/stderr, quotes structured Windows arguments, and distinguishes
  completion, timeout, cancellation, wait, capture, and launch failures. The
  `test_support_child_process_runner_contract` case covers output beyond pipe
  capacity, bounded retention, Unicode/quotes/trailing backslashes, nonzero exit,
  unrelated inheritable handles, descendant kill, cancellation, and launch
  failure. The focused Compare runner passed 1 / 0 / 0 under run id
  `20260715T173411Z-93784-77ccd2958a9645ee9839834eb972ef31`; the mandatory
  contamination-recovery full Debug/x64 rebuild was zero-warning/zero-error in
  `.build/logs/msbuild-20260715_193016_564.log`.

- 2026-07-14 RedConfigure manager closeout: Windows SDK `rc.exe`
  compilation of generated resources; language-column, batch, rectangular
  clipboard, accelerator, combined-validation, theme metadata/contrast/recipe,
  undo/redo, grouped JSON5, unbalanced-placeholder, four-page UI smoke, and
  deterministic 6-owner/1,500-row/4-culture/10-theme performance coverage.

- 2026-07-17 Observatory Track 19 RedConfigure checkpoint: all ten theme recipes and all seven localization
  batch kinds now have typed preview characterization; exact-reference, direct-source-only conversion,
  malformed/missing-target, stale-source, atomic rejection, session Undo, typed validation, duplicate-theme
  boundary/result, and comma-decimal locale contracts pass. A zero-warning Debug build and five candidate
  performance samples are archived at
  `Specs/TestRuns/4cb089111a23/RedConfigure/2026-07-17_1548_observatory_track19/`; the 512-token validated mass
  preview median is 57,111 us against its 100 ms explicit-action budget.

- 2026-07-10 test-suite stabilization final Full closeout: after the
  SearchService foreground-readiness diagnostics and FileOperations queued-cancel
  staging fixes, a clean `Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2`
  passed under the required short NTFS root `C:\RSPerf` with build enabled. Run
  id `20260710T200626Z-43532-7869ee949a30441c8d586e476fd428d3` reported
  1173 total cases, 1121 passed, 0 failed, 52 expected skips, disk audit clean
  with 0 issues, and 0 flaky/regression/isolation-suspect/unclassified-failure
  classifications. Final accepted evidence is archived at
  `Specs\TestRuns\8fa8954fe9ff\Full\2026-07-10_2247_full_green\last_run`.
  Source-contract coverage now guards foreground SearchService status waits with
  process-exit diagnostics and FileOperations queued-cancel tests with staged
  active/queued task proof rather than timing assumptions.

- 2026-07-10 test-suite stabilization Commands convergence checkpoint: added
  source-contract coverage for the Quick Search integrated navigation
  no-match/focus-reactivation path and the menu-bar keyboard-open stale-pointer
  seed used by hover-switch tests. Focused `TestHarnessSourceContracts` passed
  121 / 0 / 0, Debug `RedSalamander` build log
  `.build\logs\msbuild-20260710_192905_867.log` had 0 warnings / 0 errors,
  focused Quick Search and menu-bar proofs passed, and broad
  `Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2` passed
  786 / 0 / 2 with disk audit 0 under `C:\RSPerf`.

- 2026-07-06 test-suite stabilization Pester runner compatibility checkpoint:
  the clean `Suite Full` disk-proof run exposed that `Run-AllTests.ps1` passed
  raw `Invoke-Pester -Path ... -PassThru` argument strings, which Pester 3.4
  misbound into `EnableExit` and failed before tooling tests executed. Added
  `New-RSPesterInvokeParameters(...)` to bind `-Script` for Pester 3 and `-Path`
  for newer Pester versions, with explicit `-Tag`/`-ExcludeTag` forwarding.
  Added `RunAllTestsPlan.Tests.ps1` coverage for both parameter sets.

- 2026-07-06 test-suite stabilization stale TestSandbox run cleanup checkpoint:
  added `RunAllTestsPlan.Tests.ps1` coverage proving the runner can parse
  PID-bearing `runs\<timestamp>-<pid>-<guid>` directory names, select only
  parseable sibling run directories whose owner PID is no longer live, and
  leave the current run, explicitly allowed runs, live-PID runs, and
  unparseable/manual run directories untouched. `Run-AllTests.ps1` now sweeps
  those stale dead-PID siblings before build and child-test execution, independent
  of the legacy sandbox cleanup switch. GREEN focused check: `RunAllTestsPlan`
  32 / 0 / 0. Focused runner proof
  `Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter showIdentical
  -TimeoutMultiplier 0.1 -SkipLegacySandboxCleanup` passed Compare 1 / 0 / 0
  with run id `20260706T145934Z-78632-fe9b8b010cce4dbba8cef15d81d31009`;
  the runner removed 48 stale PID/GUID-scoped sibling run directories with 0
  cleanup failures, and the post-run audit dropped to 13 remaining non-blocking
  issues: 12 unparseable historical proof run dirs plus the known legacy
  selftest root. After deleting those local historical proof dirs and repairing
  ACLs on legacy hang-dump artifacts under `%LOCALAPPDATA%\RedSalamander\SelfTest`,
  the same focused runner command with legacy cleanup enabled passed Compare
  1 / 0 / 0 with run id
  `20260706T151836Z-38328-966786696a054e21872134aeee3860a6`; stale cleanup
  removed the prior PID/GUID run, reported 0 cleanup failures, and the aggregate
  `test_sandbox_audit` reported 0 issues.

- 2026-07-06 test-suite stabilization Compare dummy filesystem TestSandbox
  checkpoint: added `TestHarnessSourceContracts.Tests.ps1` coverage proving
  `dummy_content`, `normalized_name_collision_preserves_same_side_entries`,
  `deep_tree`, and `invalidate` acquire native Compare scratch roots through
  `SelfTest::AcquireTestSandbox(SelfTest::SelfTestSuite::CompareDirectories, ...)`
  and no longer allocate under fixed `Y:\CompareSelfTest_*`,
  `Z:\CompareSelfTest_*`, or `W:\CompareSelfTest_*` roots. RED Pester first
  failed at 81 / 1 / 0 against the hard-coded `dummy_content` root; GREEN
  focused checks: `TestHarnessSourceContracts` 82 / 0 / 0 and `TestInventory`
  5 / 0 / 0. Debug `RedSalamander` build log
  `.build\logs\msbuild-20260706_162932_943.log` had 0 warnings / 0 errors.
  Focused Compare proof passed 4 / 0 / 0 for the four migrated cases with run
  id `20260706T143152Z-74460-563d6bd4c17b4a509b9d606ccb14c2da`; the shared
  trace records all four roots under
  `.build\TestSandbox\runs\20260706T143152Z-74460-563d6bd4c17b4a509b9d606ccb14c2da\scratch\compare\<case>`.
  The broad source scan now reports only guard/cleanup literals in
  `Tools\Tests`.

- 2026-07-06 test-suite stabilization broad raw-root source-guard checkpoint:
  added `TestHarnessSourceContracts.Tests.ps1` coverage that enumerates primary
  gated test sources under `RedSalamander\SelfTest`, `Tests`, and `Tools\Tests`
  and fails on raw temp/profile/root acquisition patterns:
  `std::filesystem::temp_directory_path(...)`, `GetTempPath*`,
  `GetTempFileName*`, PowerShell `[System.IO.Path]::GetTempPath()`, direct
  `LOCALAPPDATA`/`TEMP`/`TMP` env acquisition, and known legacy root families
  such as `CompareSelfTest_*`, `RedSalamanderCrossVolumeSelfTest_*`,
  `C:\BatchRename*SelfTest`, `Specs\TestRuns\DxUiGallery`, and
  `Specs\TestRuns\local_scratch`. The guard includes synthetic violation proof
  and allows `RunAllTestsPlan.Tests.ps1` to mention historical cross-volume
  roots only as cleanup-plan evidence. GREEN focused check:
  `TestHarnessSourceContracts` 83 / 0 / 0.

- 2026-07-06 test-suite stabilization TestSandbox disk-audit checkpoint:
  added `RunAllTestsPlan.Tests.ps1` coverage for
  `Get-RSTestSandboxDiskAudit(...)`, proving clean current-run state,
  stale sibling run-directory detection, unexpected direct-child detection under
  the TestSandbox root, and resolved legacy selftest/temp/cross-volume root
  reporting. `Run-AllTests.ps1` now preserves the non-blocking
  `test_sandbox_audit` object in `run-all-tests-results.json` and prints the
  audit issue count in the artifacts section. GREEN focused check:
  `RunAllTestsPlan` 31 / 0 / 0. Focused runner proof
  `Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter showIdentical
  -TimeoutMultiplier 0.1 -SkipLegacySandboxCleanup` passed Compare 1 / 0 / 0
  with run id `20260706T145229Z-74348-9f42fae656424187a176718f128249f9`; the
  aggregate JSON preserved `test_sandbox_audit` with 60 non-blocking issues
  (`unexpected-test-run-dir=59`, `legacy-selftest-root=1`), proving the audit
  plumbing catches real stale disk state.

- 2026-07-06 test-suite stabilization legacy cleanup locked-target checkpoint:
  added `RunAllTestsPlan.Tests.ps1` coverage proving a locked historical
  `%LOCALAPPDATA%\RedSalamander\SelfTest` target is reported as `Status=Failed`
  with the removal error instead of aborting `Tools\Clean-TestSandbox.ps1`.
  Added `TestHarnessSourceContracts.Tests.ps1` coverage proving the cleanup
  script uses `-ErrorAction Stop`, catches removal failures, writes warnings,
  and preserves failed-status result rows. GREEN focused checks:
  `RunAllTestsPlan` 30 / 0 / 0 and `TestHarnessSourceContracts` 82 / 0 / 0.
  A normal focused Compare proof with legacy cleanup enabled emitted a warning
  for stale locked dump `RedSalamander-commands-find-setup-hang-p37204.dmp`,
  then continued and passed the four migrated Compare cases 4 / 0 / 0 with run
  id `20260706T143655Z-85708-3594b0f9798e45989d6c1bdbffb097cf`.

- 2026-07-06 test-suite stabilization Tools Pester TestSandbox checkpoint:
  added `RunAllTestsPlan.Tests.ps1` coverage for
  `New-RSTestSandboxScratchDirectory(...)` and
  `TestHarnessSourceContracts.Tests.ps1` coverage proving
  `WingetValidation.Tests.ps1`, `VcpkgInstallSafety.Tests.ps1`, and
  `ShowPerfRuns.Tests.ps1` no longer allocate scratch roots through
  `[System.IO.Path]::GetTempPath()`. Tooling Pester scratch now routes under
  `REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\tools-pester\<case>`, and
  legacy cleanup includes the former `RSWingetValidationTest_*`,
  `RSVcRuntimeTest_*`, `rs-vcpkg-install-root-test*`,
  `rs-vcpkg-single-file-merge-*`, and `rs-show-perfruns-tests-*` temp roots.
  RED source-contract Pester first failed at 80 / 1 / 0 against the missing
  helper/migration; GREEN focused checks: `RunAllTestsPlan` 30 / 0 / 0,
  `TestHarnessSourceContracts` 81 / 0 / 0, `WingetValidation` 17 / 0 / 0,
  `VcpkgInstallSafety` 5 / 0 / 0, `ShowPerfRuns` 5 / 0 / 0, and
  `TestInventory` 5 / 0 / 0. The raw `[System.IO.Path]::GetTempPath()` scan
  under `Tools\Tests` returned no matches.

- 2026-07-06 test-suite stabilization MTP fake-journal `LOCALAPPDATA`
  TestSandbox checkpoint: added `TestHarnessSourceContracts.Tests.ps1`
  coverage proving fake MTP overwrite-journal state now redirects
  `LOCALAPPDATA` through native
  `SelfTest::AcquireTestSandbox(SelfTest::SelfTestSuite::CompareDirectories, ...)`
  roots instead of writing under the user's
  `%LOCALAPPDATA%\RedSalamander\PluginState\FileSystemMtp`. The guard also
  locks the shortened scratch segments (`mtp_journal_temp_retry`,
  `mtp_journal_no_temp_puid`, and `mtp_overwrite_safety_matrix`) so long
  per-run sandbox paths do not push local journal helper paths past classic
  Win32 limits. RED Pester first failed at 79 / 1 / 0 against the legacy
  `GetEnvironmentVariableW(L"LOCALAPPDATA")` helper; GREEN focused checks:
  `TestHarnessSourceContracts` 80 / 0 / 0 and `TestInventory` 5 / 0 / 0.
  Debug `RedSalamander` build log `.build\logs\msbuild-20260706_161107_992.log`
  had 0 warnings / 0 errors. Focused Compare proof passed 6 / 0 / 0 for the
  six MTP overwrite-journal cases with run id
  `20260706T141313Z-74896-721c0782d2494b1f9456647a1ff0dc51`; the trace records
  `LOCALAPPDATA` redirected under
  `.build\TestSandbox\runs\20260706T141313Z-74896-721c0782d2494b1f9456647a1ff0dc51\scratch\compare\<short-mtp-segment>`.

- 2026-07-06 test-suite stabilization DxUiTests artifact-default TestSandbox
  checkpoint: added `TestHarnessSourceContracts.Tests.ps1` coverage proving
  generated DxUi artifact defaults now resolve through
  `GetDxUiTestArtifactPath(...)` under
  `REDSALAMANDER_TEST_ROOT\runs\<runId>\artifacts\dxui` instead of
  `Specs\TestRuns\DxUiGallery` or `Specs\TestRuns\local_scratch`. The guard
  covers default `ButtonContrast` PNG output plus the Animation and WindowHost
  local perf JSONL sinks while preserving explicit caller-supplied output paths.
  RED Pester first failed at 78 / 1 / 0 against the missing helper and
  repo-local `Specs\TestRuns` defaults; GREEN focused check:
  `TestHarnessSourceContracts` 79 / 0 / 0. Debug `DxUiTests` build log
  `.build\logs\msbuild-20260706_155547_780.log` had 0 warnings / 0 errors.
  Runtime proof passed `DxUiTests.exe --suite=ButtonContrast`,
  `--suite=Animation`, and `--suite=WindowHost` with
  `REDSALAMANDER_TEST_RUN_ID=dxui-artifact-proof-20260706-1557`; artifact root
  `.build\TestSandbox\runs\dxui-artifact-proof-20260706-1557\artifacts\dxui`
  contained `DxUiButtonContrast.png` (380558 bytes),
  `dxui_animation_scheduler_testlocal.jsonl` (85902 bytes), and
  `dxui_windowhost_stage_metrics_testlocal.jsonl` (6288 bytes).

- 2026-07-06 test-suite stabilization ViewerPETests TestSandbox checkpoint:
  added `TestHarnessSourceContracts.Tests.ps1` coverage proving ViewerPETests
  runtime fixture files now acquire standalone scratch paths through
  `AcquireViewerPETestSandbox(...)` under
  `REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\viewer-pe\<case>` and no longer
  create `ViewerWebTests`, `ViewerImgRawTests`, `ViewerImgRawPngTests`,
  `ViewerTextTests`, `ViewerSpaceTests`, or `ViewerVLCTests` folders in the
  build output directory. RED Pester first failed at 77 / 1 / 0 against the
  missing helper and build-dir fixture roots; GREEN focused check:
  `TestHarnessSourceContracts` 78 / 0 / 0 and `TestInventory` 5 / 0 / 0.
  Debug `ViewerPETests` build log `.build\logs\msbuild-20260706_154301_566.log`
  had 0 warnings / 0 errors. Focused runtime proof passed
  `TestViewerImgRawDecodesPngThroughWicWithoutErrorAlert` with
  `REDSALAMANDER_TEST_RUN_ID=viewer-pe-sandbox-proof-20260706-1543`; the case
  root
  `.build\TestSandbox\runs\viewer-pe-sandbox-proof-20260706-1543\scratch\viewer-pe\viewerimgraw_png_decode`
  existed with child_count=0 after cleanup.

- 2026-07-06 test-suite stabilization BatchRename TestSandbox checkpoint:
  added `TestHarnessSourceContracts.Tests.ps1` coverage proving BatchRename
  window fixture roots now acquire native Commands scratch paths through
  `SelfTest::AcquireTestSandbox(SelfTestSuite::Commands, ...)` and no longer
  use fixed `C:\BatchRename*SelfTest` roots. RED Pester first failed at
  76 / 1 / 0 against the missing helper and hardcoded roots; GREEN focused
  checks: `TestHarnessSourceContracts` 77 / 0 / 0 and `TestInventory` 5 / 0 / 0.
  Debug `RedSalamander` build log `.build\logs\msbuild-20260706_153105_584.log`
  had 0 warnings / 0 errors. Focused Commands proof passed 14 / 0 / 0 for the
  migrated BatchRename window-root cases with run id
  `20260706T133314Z-83440-414bdc55817245afa50b7b5d94b1fc91`, archived at
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-06_153316`.

- 2026-07-06 test-suite stabilization DxUi focus/foreground gating checkpoint:
  added `TestHarnessSourceContracts.Tests.ps1` coverage proving shared DxUi test
  helpers guard focus-sensitive Win32 assertions, `NativeTextInput` focus/caret
  paths use `TryFocusDxUiTestWindow`, Menu foreground popup/focus paths use
  `TryActivateDxUiTestWindow`, and direct discarded
  `SetForegroundWindow(ownerWindow.Hwnd())` calls are gone. The final tightened
  RED Pester guard first failed at 75 / 1 / 0 against the missing focused
  helper split; GREEN focused checks: `TestHarnessSourceContracts` 76 / 0 / 0
  and `TestInventory` 5 / 0 / 0. Debug `DxUiTests` build log
  `.build\logs\msbuild-20260706_152402_723.log` had 0 warnings / 0 errors.
  Runtime `DxUiTests.exe --suite=NativeTextInput` passed with no focus-gate
  skips in this session. Runtime `DxUiTests.exe --suite=Menu` passed while
  foreground-only cases emitted explicit `SKIPPED:` reasons in this background
  desktop session.

- 2026-07-06 test-suite stabilization PerformanceTests2 TestSandbox checkpoint:
  added `TestHarnessSourceContracts.Tests.ps1` coverage proving the standalone
  CppUnitTest perf harness now resolves its icon-enumeration and duplicate-path
  refresh fixture roots through `AcquirePerformanceTestSandbox(...)` under
  `<repoRoot>\.build\TestSandbox\runs\<runId>\scratch\performance-tests2\<case>`
  and no longer calls `std::filesystem::temp_directory_path`. RED Pester first
  failed at 74 / 1 / 0 against the missing helper; GREEN focused check:
  `TestHarnessSourceContracts` 75 / 0 / 0. Debug `PerformanceTests2` build log
  `.build\logs\msbuild-20260706_150744_273.log` had 0 warnings / 0 errors.
  Direct VSTest proof passed 12 / 0 / 0 with
  `REDSALAMANDER_TEST_RUN_ID=performance-tests2-sandbox-proof-20260706-1508`,
  and the `scratch\performance-tests2` parent had no child fixtures after cleanup.
  A source scan for `std::filesystem::temp_directory_path`, `GetTempPathW`, and
  `GetTempFileNameW` under `Tests` and `RedSalamander\SelfTest` returned no
  remaining matches.

- 2026-07-06 test-suite stabilization ShellCommands shortcut-save TestSandbox checkpoint:
  added `TestHarnessSourceContracts.Tests.ps1` coverage proving the long-path
  `SaveShellLinkForShellCommandTest(...)` temp-save fallback now acquires a
  native scratch root through
  `SelfTest::AcquireTestSandbox(SelfTest::SelfTestSuite::Commands, L"shell_shortcut_save_temp")`
  and no longer calls `GetTempPathW` or `GetTempFileNameW`. RED Pester first
  failed at 73 / 1 / 0 against the Win32 process-temp fallback; GREEN focused
  check: `TestHarnessSourceContracts` 74 / 0 / 0. Debug `RedSalamander` build
  log `.build\logs\msbuild-20260706_145453_462.log` had 0 warnings / 0 errors.
  Focused ShellCommands shortcut-helper coverage passed 4 / 0 / 0 for the
  `.lnk` go-to-target and item-properties cases with run id
  `20260706T125718Z-73056-35658bb24a43400da7d6788684b726bb`, archived at
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-06_145720`.

- 2026-07-06 test-suite stabilization Commands plugin-config TestSandbox checkpoint:
  added `TestHarnessSourceContracts.Tests.ps1` coverage proving the ViewerText
  perf fixtures and ViewerImgRaw/ViewerWeb close-roundtrip fixtures now acquire
  native scratch roots through
  `SelfTest::AcquireTestSandbox(SelfTest::SelfTestSuite::Commands, ...)` under
  `<repoRoot>\.build\TestSandbox\runs\<runId>\scratch\commands\<case>`, and no
  longer call `std::filesystem::temp_directory_path`. RED Pester first failed
  at 72 / 1 / 0 against the old process-temp roots; GREEN focused check:
  `TestHarnessSourceContracts` 73 / 0 / 0. Debug `RedSalamander` build log
  `.build\logs\msbuild-20260706_144915_221.log` had 0 warnings / 0 errors.
  A focused Commands proof passed 4 / 0 / 0 for
  `viewer_text_hex_byte_color_perf`, `viewer_text_diff_perf`,
  `viewer_imgraw_close_roundtrip`, and `viewer_web_close_roundtrip` with run id
  `20260706T125122Z-78724-35dc0b8689764f4eaf08d9357f6d1939`, archived at
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-06_145125`; the trace records
  all four `TestSandbox: suite=commands` case roots and the
  `scratch\commands` directory had no remaining child fixtures after cleanup.

- 2026-07-06 test-suite stabilization CrashHandling TestSandbox checkpoint:
  added `TestHarnessSourceContracts.Tests.ps1` coverage proving
  CrashHandlingTests marker-file scratch roots now derive from
  `REDSALAMANDER_TEST_ROOT` + `REDSALAMANDER_TEST_RUN_ID` under
  `<repoRoot>\.build\TestSandbox\runs\<runId>\scratch\crash-handling\marker-files`,
  and no longer call `std::filesystem::temp_directory_path`. RED Pester first
  failed at 71 / 1 / 0 against the missing helper; GREEN focused check:
  `TestHarnessSourceContracts` 72 / 0 / 0. A focused Debug
  `CrashHandlingTests.exe` proof passed with run id
  `crash-handling-sandbox-proof-20260706-1441` and left no child fixtures under
  the `crash-handling` scratch root.

- 2026-07-06 test-suite stabilization ViewerSqlite TestSandbox checkpoint:
  added `TestHarnessSourceContracts.Tests.ps1` coverage proving ViewerSqliteTests
  database scratch roots now derive from `REDSALAMANDER_TEST_ROOT` +
  `REDSALAMANDER_TEST_RUN_ID` under
  `<repoRoot>\.build\TestSandbox\runs\<runId>\scratch\viewer-sqlite\database`,
  and no longer call `std::filesystem::temp_directory_path`. RED Pester first
  failed at 70 / 1 / 0 against the missing helper; GREEN focused check:
  `TestHarnessSourceContracts` 71 / 0 / 0. A focused Debug
  `ViewerSqliteTests.exe` proof passed with run id
  `viewer-sqlite-sandbox-proof-20260706-1439` and left no child fixtures under
  the `viewer-sqlite` scratch root.

- 2026-07-06 test-suite stabilization RedConfigure TestSandbox checkpoint:
  added `TestHarnessSourceContracts.Tests.ps1` coverage proving RedConfigureTests
  scratch roots now derive from `REDSALAMANDER_TEST_ROOT` +
  `REDSALAMANDER_TEST_RUN_ID` under
  `<repoRoot>\.build\TestSandbox\runs\<runId>\scratch\redconfigure\<case>`,
  and no longer call `std::filesystem::temp_directory_path`. RED Pester first
  failed at 69 / 1 / 0 against the missing helper; GREEN focused check:
  `TestHarnessSourceContracts` 70 / 0 / 0. A focused Debug
  `RedConfigureTests.exe` proof passed 23 / 0 / 0 with run id
  `redconfigure-sandbox-proof-20260706-1431`.

- 2026-07-06 test-suite stabilization alternate-volume TestSandbox checkpoint:
  added `TestHarnessSourceContracts.Tests.ps1` coverage proving FileOperations
  real cross-volume scratch allocates through
  `SelfTest::AcquireTestSandboxOnVolume(...)` under
  `<AltDrive>:\RedSalamanderTestSandbox\runs\<runId>\scratch\fileops\real_cross_volume_move`,
  prunes empty alternate-volume run parents after cleanup, and no longer creates
  `RedSalamanderCrossVolumeSelfTest_*` roots. RED Pester first failed at
  68 / 1 / 0 against the missing pruning helper; GREEN focused check:
  `TestHarnessSourceContracts` 69 / 0 / 0.

- 2026-07-06 test-suite stabilization raw wait scaling checkpoint:
  added `TestHarnessSourceContracts.Tests.ps1` coverage proving the named
  raw waits from the flake-convergence plan are now multiplier-scaled:
  Search `WaitForFlag(...)` scales internally, BatchRename's gated
  `WaitForSingleObject(...)` uses `SelfTest::ScaleTimeout(30'000u)`, and
  FileOperations directory recreation uses a scaled 6-second deadline with
  50 ms retry slices instead of a fixed attempt count. RED Pester first
  failed at 67 / 1 / 0 against the unscaled waits; GREEN focused check:
  `TestHarnessSourceContracts` 68 / 0 / 0.

- 2026-07-06 test-suite stabilization FolderView overlay perf advisory checkpoint:
  added `TestHarnessSourceContracts.Tests.ps1` coverage proving
  `folderView_perf_overlay_invalidation_stress` no longer uses the fixed
  `Scale(4200ms)` collection window, does not hard-fail on
  `overlaySamplesEnoughForP95`, still writes `metricQuality`, and records the
  overlay frame collection pass count. GREEN focused check:
  `TestHarnessSourceContracts` 67 / 0 / 0.

- 2026-07-06 test-suite stabilization foreground search-service JobObject checkpoint:
  added `TestHarnessSourceContracts.Tests.ps1` coverage proving the Compare
  foreground search-service selftest harness creates a kill-on-close JobObject,
  starts the service suspended, assigns it with `AssignProcessToJobObject`,
  and resumes only after assignment succeeds. GREEN focused check:
  `TestHarnessSourceContracts` 66 / 0 / 0.

- 2026-07-06 test-suite stabilization native TestSandbox scratch checkpoint:
  added `TestHarnessSourceContracts.Tests.ps1` coverage proving native
  selftests expose `SelfTest::AcquireTestSandbox`, route unified-run scratch
  under `runs\<runId>\scratch\<suite>\<case>`, keep scratch separate from
  `last_run` artifacts, and move Compare foreground search-service stdout
  capture off process temp. GREEN focused check: `TestHarnessSourceContracts`
  65 / 0 / 0.

- 2026-07-06 test-suite stabilization local index snapshot reload timing checkpoint:
  added `TestHarnessSourceContracts.Tests.ps1` coverage proving
  `local_index_core_snapshot_reload` emits
  `compare.selftest.local_index.snapshot_reload_us` and no longer gates correctness on
  `warmElapsedMs < 1000u`. GREEN focused check: `TestHarnessSourceContracts` 63 / 0 / 0.

- 2026-07-06 test-suite stabilization native TestSandbox root checkpoint:
  added `TestHarnessSourceContracts.Tests.ps1` coverage proving native
  selftests consume `REDSALAMANDER_TEST_ROOT` plus `REDSALAMANDER_TEST_RUN_ID`
  directly, derive `runs\<runId>\artifacts\selftest`, and that the normal
  runner clears inherited `REDSALAMANDER_SELFTEST_ROOT` values before child
  execution. The legacy override remains available only for explicit
  compatibility launches such as the reviewed quarantine repair lane. GREEN
  focused check: `TestHarnessSourceContracts` 62 / 0 / 0. GREEN Debug build
  `.build\logs\msbuild-20260706_125917_741.log` had 0 warnings / 0 errors.
  Runtime proof with `REDSALAMANDER_SELFTEST_ROOT` intentionally set to
  `Z:\src\RedSalamander\.build\BogusLegacySelfTestRoot` passed
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter showIdentical -TimeoutMultiplier 0.1 -SkipLegacySandboxCleanup`
  with 1 passed / 0 failed under
  `.build\TestSandbox\runs\20260706T110346Z-87828-dab956516dda4f7ca2c88af9da3d1580\artifacts\selftest\last_run`,
  proving the normal runner ignored the inherited legacy root and the native suite wrote directly to
  the runner-owned root.

- 2026-07-06 test-suite stabilization legacy TestSandbox cleanup checkpoint:
  added `RunAllTestsPlan.Tests.ps1` coverage proving the legacy cleanup plan targets
  `%LOCALAPPDATA%\RedSalamander\SelfTest`, known `%TEMP%` standalone/perf roots, and
  fixed-drive `RedSalamanderCrossVolumeSelfTest_*` roots without resolving wildcards during
  planning. Added `TestHarnessSourceContracts.Tests.ps1` coverage proving
  `Tools\Clean-TestSandbox.ps1` stays dry-run gated, requires `-Apply`, uses `ShouldProcess`,
  deletes only via `Remove-Item -LiteralPath`, and is invoked by `Run-AllTests.ps1` before child
  tests unless `-SkipLegacySandboxCleanup` is supplied. GREEN focused checks: `RunAllTestsPlan`
  29 / 0 / 0, `TestHarnessSourceContracts` 61 / 0 / 0, dry-run cleanup against nonexistent
  controlled roots removed nothing, and parser validation passed for `Run-AllTests.ps1`,
  `TestRunPlan.ps1`, and `Clean-TestSandbox.ps1`.

- 2026-07-06 test-suite stabilization runtime classifier-proof checkpoint:
  added `RunAllTestsPlan.Tests.ps1` coverage proving explicit
  `-SelfTestFlakyProofCase` / `-SelfTestOrderProofCase` forwarding and focused
  `-CaseFilter` shuffle-triage planning. Added
  `TestHarnessSourceContracts.Tests.ps1` coverage proving native debug proof
  hooks, classifier proof suite/shuffle context, and explicit-order dispatch
  wiring for Commands, CompareDirectories, and FileOperations. Runtime proof:
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter 'unique,showIdentical' -ClassifyFailures -SelfTestFlakyProofCase showIdentical -TimeoutMultiplier 0.1`
  exited `1` as expected and classified
  `20260706T104131Z-85784-c110c5635f9b4b9886e8f2432b8d4f1b`
  as blocking `FLAKY` with retry modes
  `failed-case,shuffle-triage,shuffle-triage,shuffle-triage`, shuffle seeds
  `885711464,370821081,1015852695`, and counts
  `flaky=1 regression=0 isolation=0`. The order-dependent proof command
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter 'unique,showIdentical' -ClassifyFailures -SelfTestOrderProofCase showIdentical -TimeoutMultiplier 0.1`
  exited `1` as expected and classified
  `20260706T104138Z-85864-cecd27054aed4e3387c8d36c284b1f50`
  as blocking `REGRESSION` with three failing shuffle-triage attempts and counts
  `flaky=0 regression=1 isolation=0`.

- 2026-07-06 test-suite stabilization Commands repeat/shuffle checkpoint:
  added `RunAllTestsPlan.Tests.ps1` coverage proving native self-test repeat
  and shuffle arguments are forwarded, fractional timeout multipliers use
  invariant decimal formatting, repeat result rows do not fail result-coverage
  validation as duplicate drift, and shuffle seeds are forwarded to all three
  in-product suites after their explicit-order migrations. Added
  `TestHarnessSourceContracts.Tests.ps1` coverage proving native repeat/shuffle
  CLI parsing, shared execution-order helpers, Commands explicit seeded order,
  FileOperations native repeat aggregation and explicit phase-order seeded shuffle, and
  CompareDirectories explicit-order seeded shuffle. Runtime proof:
  `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter modeless_window_ownership -SelfTestRepeat 2 -SelfTestShuffleSeed 123 -TimeoutMultiplier 0.1`
  passed with 2 cases, `repeat_index` values `1,2`, `repeat_count=2`, and
  `shuffle_seed=123` in both suite and aggregate JSON under
  `.build\TestSandbox\runs\20260706T094019Z-77120-cb4f59ca0800460c869afadf964b779e\artifacts\selftest\last_run`.
  `.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter FileOps_ProviderCapabilityMatrix -SelfTestRepeat 2 -SelfTestShuffleSeed 789 -TimeoutMultiplier 0.1`
  passed with 6 FileOperations rows across `repeat_index` values `1,1,1,2,2,2`,
  `repeat_count=2`, `shuffle_seed=789`, and trace line `FileOpsSelfTest: explicit execution order`
  under `.build\TestSandbox\runs\20260706T100312Z-74672-1393b09f8812458ba12f7890e2e352ca\artifacts\selftest\last_run`.
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter showIdentical -SelfTestRepeat 2 -SelfTestShuffleSeed 456 -TimeoutMultiplier 0.1`
  passed with 2 CompareDirectories rows, `repeat_index` values `1,2`, `repeat_count=2`, and
  `shuffle_seed=456` under
  `.build\TestSandbox\runs\20260706T094952Z-52548-5efbb28842ce47708551b52945b161da\artifacts\selftest\last_run`.
  The same slice fixed runner-native case listing so `--selftest-list-cases`
  cannot overwrite prior run artifacts.

- 2026-07-06 test-suite stabilization partial crash-result checkpoint: added
  `TestHarnessSourceContracts.Tests.ps1` coverage proving `RunCase` records the
  current in-flight case, shared case-result emission flushes suite
  `results.json` after every case, the self-test SEH helper marks the in-flight
  case `crashed` and writes partial aggregate `results.json`, and
  `Run-AllTests.ps1` treats `crashed` as red failure evidence. RED Pester failed
  at 53 / 1 / 0 against the missing hooks; GREEN Pester passed 54 / 0 / 0 after
  adding the common in-flight tracker, crashed status, partial JSON writer, and
  runner parser handling. GREEN Debug build passed with
  `.build/logs/msbuild-20260706_105446_011.log`. The deterministic crash proof
  uses `--selftest-crash-case=NAME`; `modeless_window_ownership` exited `-1` in
  915 ms and archived
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-06_110551/selftest_run_results.json`
  with status `crashed` and reason `0xC0000005` Access Violation.

- 2026-07-06 test-suite stabilization fatal-modal bypass checkpoint: added
  `TestHarnessSourceContracts.Tests.ps1` coverage proving
  `ShowFatalErrorDialog(...)` checks `IsRunningAnySelfTest()`, appends a
  self-test trace row, and returns before `dialog.ShowModal()` so fatal
  self-test paths can exit non-zero without blocking CI. RED Pester failed at
  52 / 1 / 0 against the unconditional modal; GREEN Pester passed 53 / 0 / 0
  after adding the guarded bypass in `RedSalamander.cpp`.

- 2026-07-06 test-suite stabilization quarantine repair-lane checkpoint:
  added `RunAllTestsPlan.Tests.ps1` coverage proving reviewed quarantine
  entries are matched against actual runner harness adapter names, self-test
  entries must exist in the adapter case list before they become single-case
  repair attempts with `--selftest-case=<name>`, unknown harnesses are invalid
  blocking entries, and
  `run-all-tests-results.json` preserves quarantine repair attempts and
  reproduced counts. RED Pester first failed at 15 / 3 / 0 against the missing
  repair-plan helper and summary parameter; GREEN Pester passed 18 / 0 / 0
  after adding `Get-RSTestQuarantineRepairPlan`, repair-lane execution in
  `Run-AllTests.ps1`, self-test case-list validation for active quarantines,
  and aggregate summary fields.

- 2026-07-06 test-suite stabilization GitHub summary checkpoint: added
  `RunAllTestsPlan.Tests.ps1` coverage proving the runner can format a GitHub
  step summary with suite status, blocking classification counts, active
  quarantine owner/expiry, and repair-lane pass/fail evidence. RED Pester
  failed at 19 / 1 / 0 against the missing formatter; GREEN Pester passed
  20 / 0 / 0 after adding `Convert-RSTestRunSummaryToGitHubStepSummary` and
  appending it to `GITHUB_STEP_SUMMARY` when GitHub Actions provides that file.

- 2026-07-05 GR-12 IconCache path failure-store race checkpoint: added
  `TestHarnessSourceContracts.Tests.ps1` coverage proving path failure stores
  use insert-only semantics, never downgrade a concurrent successful icon-index
  entry to `lookupFailed=true`, and emit `iconcache.duplicate_path_query_race`
  when another thread wins the store race. RED Pester failed at 51 / 1 / 0
  against the `insert_or_assign(...)` failure store; GREEN Pester passed
  52 / 0 / 0 after changing the store to `emplace(...)`.

- 2026-07-05 GR-C6 FileSystemMtp warning-suppression rationale checkpoint:
  added `TestHarnessSourceContracts.Tests.ps1` coverage proving each
  FileSystemMtp `#pragma warning(disable: ...)` site named in Granite carries
  a local rationale comment matching the sibling plugin patterns for WIL and
  yyjson warning noise. RED Pester failed at 50 / 1 / 0 against the uncommented
  MTP suppressions; GREEN Pester passed 51 / 0 / 0 after adding comments.
  GREEN plugin build `.build\logs\msbuild-20260705_183102_114.log` produced
  `FileSystemMtp.dll` with 0 warnings / 0 errors.

- 2026-07-05 GR-C5 DxUi NativeTextInput formatting checkpoint: added
  `TestHarnessSourceContracts.Tests.ps1` coverage proving
  `TraceNativeTextInputTsfStep(...)` uses `std::array<wchar_t, 768>` plus
  `std::format_to_n(...)` bounded formatting and no longer uses
  `wchar_t line[768]` / `StringCchPrintfW`. RED Pester failed at
  49 / 1 / 0 against the old formatting block; GREEN Pester passed 50 / 0 / 0
  after conversion. GREEN app build
  `.build\logs\msbuild-20260705_182416_209.log` produced the Debug app with
  0 warnings / 0 errors, and GREEN `DxUiTests` build
  `.build\logs\msbuild-20260705_182634_951.log` produced the Debug test binary
  with 0 warnings / 0 errors. Focused `DxUiTests.exe --suite=NativeTextInput`
  passed; perf artifact
  `Specs\TestRuns\local_scratch\dxui_native_textinput_gr_c5_format_to_n_20260705_1828.jsonl`.

- 2026-07-05 GR-C4 SearchAndIndex callback exception-handling checkpoint:
  added `TestHarnessSourceContracts.Tests.ps1` coverage proving the
  `sqlite_index_store_load_and_apply_journal_delta` and
  `sqlite_index_store_root_lookup_case_insensitive` `EnumerateVolume(...)`
  callbacks preserve fatal `std::bad_alloc` handling and log a
  `Debug::Error(...)` before translating non-allocation `std::exception`
  failures to `E_FAIL` at their `noexcept` callback boundaries. RED Pester
  failed at 48 / 1 / 0 against the unfixed callback blocks; GREEN Pester
  passed 49 / 0 / 0 after adding the mandatory comments and logs. GREEN app
  build `.build\logs\msbuild-20260705_181849_764.log` produced the Debug app
  with 0 warnings / 0 errors. Focused Compare archives
  `Specs\TestRuns\4cb089111a23\CompareDirectories\2026-07-05_182054\` and
  `Specs\TestRuns\4cb089111a23\CompareDirectories\2026-07-05_182059\` passed
  `sqlite_index_store_load_and_apply_journal_delta` and
  `sqlite_index_store_root_lookup_case_insensitive`, respectively, each at
  1 / 0 / 0.

- 2026-07-05 GR-C3 LocalSearch snapshot exception-handling checkpoint: added
  `TestHarnessSourceContracts.Tests.ps1` coverage proving
  `OpenSnapshotTempFile(...)` preserves fatal `std::bad_alloc` handling and
  logs a `Debug::Error(...)` before translating non-allocation
  `std::exception` failures to `E_FAIL` at the `noexcept` boundary. RED Pester
  failed at 47 / 1 / 0 against the unfixed catch block; GREEN Pester passed
  48 / 0 / 0 after adding the mandatory comment and log. GREEN app build
  `.build\logs\msbuild-20260705_181218_425.log` produced the Debug app with
  0 warnings / 0 errors. Focused Compare archives
  `Specs\TestRuns\4cb089111a23\CompareDirectories\2026-07-05_181423\` and
  `Specs\TestRuns\4cb089111a23\CompareDirectories\2026-07-05_181429\` passed
  `search_source_allocation_and_folding_guard` and
  `search_low_hardening_smoke`, respectively, each at 1 / 0 / 0.

- 2026-07-05 GR-C1/GR-C2 MTP RAII ownership checkpoint: added
  `TestHarnessSourceContracts.Tests.ps1` coverage proving MTP public reader and
  writer helper objects keep their `FileSystemMtp` owner through
  `wil::com_ptr<FileSystemMtp>` without owner-local manual `AddRef`/`Release`,
  and `ReadDirectoryInfo(...)` keeps `FilesInformationMtp` in
  `std::unique_ptr` ownership until successful `IFilesInformation**` handoff.
  RED Pester first failed at 46 / 1 / 0 while `MtpBufferedWriter` still owned a
  raw `FileSystemMtp*`; GREEN Pester passed 47 / 0 / 0 after the writer,
  adjacent reader, and directory-info handoff were converted. GREEN
  `.\build.ps1 -ProjectName FileSystemMtp -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors; log
  `.build\logs\msbuild-20260705_180344_864.log`. Focused Compare archives
  `Specs\TestRuns\4cb089111a23\CompareDirectories\2026-07-05_180425\` and
  `Specs\TestRuns\4cb089111a23\CompareDirectories\2026-07-05_180431\` passed
  `mtp_fake_backend_enumerate_read_and_capabilities` and
  `mtp_public_writer_stages_until_commit`, respectively, each at 1 / 0 / 0.

- 2026-07-05 GR-S4(f) FolderView WarpDrive cleanup checkpoint: added
  `TestHarnessSourceContracts.Tests.ps1` coverage proving shell thumbnail stat
  increments route through one `IncrementThumbnailStat(...)` helper and
  FolderView uses one generic `PendingToPaintMetric` shape for both
  input-to-paint and refresh-to-paint pending metrics. RED Pester first failed
  at 44 / 1 / 0 before `IncrementThumbnailStat(...)` existed, then at
  45 / 1 / 0 before `PendingToPaintMetric` existed; GREEN Pester passed
  46 / 0 / 0 after both extractions. GREEN app build
  `.build\logs\msbuild-20260705_175607_052.log` produced the Debug app with
  0 warnings / 0 errors. Focused Commands archives
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_175832\`,
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_175839\`,
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_175845\`, and
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_175904\` passed
  `folderView_thumbnail_cached_only_no_close_stall`,
  `folderView_perf_slow_virtual_provider`,
  `folderView_perf_refresh_preservation`, and
  `folderView_perf_scroll_render_stress`, respectively, each at 1 / 0 / 0.
  The `D3D11CreateDevice` duplication subitem was already covered by the
  GR-A6 shared DxUi device-creation helper closeout.

- 2026-07-05 GR-S4(d) FolderView owned-threadpool submit helper checkpoint:
  added `TestHarnessSourceContracts.Tests.ps1` coverage proving
  `Common\Helpers.h` owns `SubmitOwnedThreadpoolCallback(...)`, the helper uses
  `TrySubmitThreadpoolCallback` and calls `owned->Execute()`, and both
  `PasteShortcutWork` and `ProviderAllowedWork` route through the helper instead
  of carrying local `TrySubmitThreadpoolCallback` / `work.release()` submit
  scaffolding. RED Pester first failed at 43 / 1 / 0 before the helper existed;
  GREEN Pester passed 44 / 0 / 0 after the extraction and test-anchor
  correction. GREEN app build `.build\logs\msbuild-20260705_174449_942.log`
  produced the Debug app with 0 warnings / 0 errors. Focused Commands archives
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_174925\` and
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_174931\` passed
  `cmd_pane_clipboardPasteShortcut_returns_before_worker_complete` and
  `folderView_perf_slow_virtual_provider`, respectively, each at 1 / 0 / 0.

- 2026-07-05 GR-S4(a) FileOperations issues-pane focus-restore helper
  checkpoint: added `TestHarnessSourceContracts.Tests.ps1` coverage proving
  `FolderWindow.FileOperationsInternal.h` owns
  `RestoreActivePaneFolderViewFocusIfWindowHadFocusBeforeHide(...)`, and both
  `ToggleIssuesPane()` and `FileOperationsIssuesPaneState::OnClose(...)` call
  it instead of carrying local folder-view focus fallback chains. RED Pester
  first failed at 42 / 1 / 0 before the helper existed; GREEN Pester passed
  43 / 0 / 0 after the extraction. GREEN app build
  `.build\logs\msbuild-20260705_173537_689.log` produced the Debug app with
  0 warnings / 0 errors. Focused Commands archive
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_173744\` passed
  `cmd_pane_fileops_issues_pane_hide_restores_folder_focus` at 1 / 0 / 0.

- 2026-07-05 GR-S4(e) ViewerSpace env-flag helper checkpoint: added
  `TestHarnessSourceContracts.Tests.ps1` coverage proving
  `Commands.SelfTest.ViewCommands.cpp` has one
  `IsViewerSpaceEnvFlagEnabled(...)` parser for `1` / `true` / `yes`, and both
  ViewerSpace opt-in scenarios call it instead of carrying local
  `GetEnvironmentVariableW(...)` readers. RED Pester first failed at 41 / 1 / 0
  before the helper existed; GREEN Pester passed 42 / 0 / 0 after the extraction.
  GREEN app build `.build\logs\msbuild-20260705_172622_129.log` produced the
  Debug app with 0 warnings / 0 errors. Focused Commands archives
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_172847\` and
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_172852\` exercised the
  normal opt-in skip path for `cmd_viewer_space_layout_20k_visible_optin` and
  `viewer_space_perf_large_optin`, respectively, each at 0 / 0 / 1.

- 2026-07-05 GR-S4(c) FileOperations env-helper reuse checkpoint: added
  `TestHarnessSourceContracts.Tests.ps1` coverage proving
  `MaybeInjectBridgeCreateDirectoryRaceForSelfTest(...)` routes its race-path
  env read through `TryReadEnvironmentVariableForSelfTest(...)` and no longer
  carries raw `GetEnvironmentVariableW` calls in that function body. RED Pester
  first failed at 40 / 1 / 0 before the helper was used; GREEN Pester passed
  41 / 0 / 0 after the function was patched. GREEN app build
  `.build\logs\msbuild-20260705_172008_860.log` produced the Debug app with 0
  warnings / 0 errors. Focused FileOps archive
  `Specs\TestRuns\4cb089111a23\FileOps\2026-07-05_172218\` passed
  `Riptide_BridgeCreateDirectoryRaceExistingFilePromptsPartial` at 3 / 0 / 0.

- 2026-07-05 GR-S4(b) shared FolderView env-flag helper checkpoint: added
  `TestHarnessSourceContracts.Tests.ps1` coverage proving
  `Common\Helpers.h` owns `EnvironmentVariables::IsTruthyFlagSet(...)`, the
  local `IsTruthySelfTestEnvironmentVariable(...)` clones are gone from
  `FolderView.Rendering.cpp` and `Commands.SelfTest.ViewCommands.cpp`, and the
  FolderView huge-perf / force-WARP gates route through the shared helper. RED
  Pester first failed at 39 / 1 / 0 before the helper existed; GREEN Pester
  passed 40 / 0 / 0 after the extraction. GREEN app build
  `.build\logs\msbuild-20260705_171253_615.log` produced the Debug app with 0
  warnings / 0 errors. Focused Commands archive
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_171549\` passed
  `folderView_perf_huge_folder_scale` with `REDSALAMANDER_FOLDERVIEW_FORCE_WARP=1`
  at 1 / 0 / 0, and its perf artifact recorded `warpRunExecuted: true`.

- 2026-07-05 GR-S2 shared DxUi modal loop checkpoint: added
  `TestHarnessSourceContracts.Tests.ps1` coverage proving `DxUi` exposes
  `RunDxUiModalLoop(...)`, archive pack/unpack prompts route their nested
  message pumps through it, those prompt blocks no longer own local
  `GetMessageW(&msg, ...)` loops, and the shared helper reposts `WM_QUIT` after
  observing it. RED Pester first failed at 38 / 1 / 0 before
  `DxUiModalLoopOptions` existed; GREEN Pester passed 39 / 0 / 0 after the
  helper and archive prompt migration. A second RED failed at 38 / 1 / 0 before
  the helper reposted `WM_QUIT`; GREEN Pester passed 39 / 0 / 0 after
  `PostQuitMessage(static_cast<int>(msg.wParam))` was added. The first focused
  pack runtime exposed the quit-propagation risk and was preserved under
  `Specs\TestRuns\4cb089111a23\Continuation\2026-07-05_1700_gr-s2_archive_prompt_hang\`.
  GREEN app build `.build\logs\msbuild-20260705_170405_259.log` produced the
  Debug app with 0 warnings / 0 errors. Focused Commands runtime archives
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_170612\` and
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_170618\` passed
  `cmd_pane_pack_prompt_uses_dxui_unique_archive_and_packer_extensions` and
  `cmd_pane_unpack_prompt_uses_dxui_destination_unpacker_and_mask`, respectively,
  each at 1 / 0 / 0.

- 2026-07-05 GR-A4 shared ordinal string helper closeout checkpoint: added
  `TestHarnessSourceContracts.Tests.ps1` coverage proving the targeted local
  case-folding predicates are gone from `ConnectionManagerWindow.cpp`,
  `FileSystemMtp.Device.cpp`, `FileSystemMtp.Shared.cpp`, and
  `SearchServiceBroker.cpp`, and that the remaining intentionally
  case-insensitive MTP/search paths use shared `OrdinalString` helpers. RED
  Pester first failed at 37 / 1 / 0 while `EqualsOrdinalIgnoreCase` still
  existed; a second RED failed at 37 / 1 / 0 while `CaseFoldKey` still used a
  local `::towlower` loop. GREEN Pester passed 38 / 0 / 0 after
  `OrdinalString::EqualsNoCase`, `OrdinalString::StartsWithNoCase`, and
  `OrdinalString::FoldCaseInvariant` were wired. GREEN app build
  `.build\logs\msbuild-20260705_164220_979.log` produced the Debug app with 0
  warnings / 0 errors. Focused Compare runtime archives
  `Specs\TestRuns\4cb089111a23\CompareDirectories\2026-07-05_164425\` and
  `Specs\TestRuns\4cb089111a23\CompareDirectories\2026-07-05_164455\` passed
  `mtp_wpd_session_and_path_cache_reuse` and
  `search_service_rejects_device_root_and_continues`, respectively, each at
  1 / 0 / 0.

- 2026-07-05 FileOperations GR-A5 bridge IO decorator ratchet checkpoint:
  added `TestHarnessSourceContracts.Tests.ps1` coverage proving the cross-FS
  bridge source-size fault injection routes through test-only
  `IFileSystemIO`/`IFileReader` decorators at bridge IO acquisition, and no
  longer through inline `CopyFileWithBuffer(...)` logic. RED Pester first failed
  at 36 / 1 / 0 because `enum class SelfTestBridgeIoRole` did not exist; GREEN
  Pester passed 37 / 0 / 0 after the decorator seam was added. GREEN app build
  `.build\logs\msbuild-20260705_162741_277.log` produced
  `.build\x64\Debug\RedSalamander.exe` with 0 warnings / 0 errors. Focused
  FileOps runtime archives
  `Specs\TestRuns\4cb089111a23\FileOps\2026-07-05_162948\` and
  `Specs\TestRuns\4cb089111a23\FileOps\2026-07-05_163017\` passed
  `Floodgate_CrossFsCopyGetSizeFailureRefusesCommit` and
  `Floodgate_CrossFsMoveGetSizeFailurePreservesSource`, respectively, each at
  3 / 0 / 0.

- 2026-07-05 FileOperations GR-S1/GR-A5 pause-point ratchet checkpoint:
  added `TestHarnessSourceContracts.Tests.ps1` coverage proving the duplicated
  post-finished-completion and bridge-move-source-cleanup self-test pauses flow
  through one `SelfTestPausePoint` helper with two instances. RED Pester first
  failed at 35 / 1 / 0 because `struct SelfTestPausePoint final` did not exist;
  GREEN Pester passed 36 / 0 / 0 after the extraction. GREEN app build
  `.build\logs\msbuild-20260705_161428_240.log` produced
  `.build\x64\Debug\RedSalamander.exe` with 0 warnings / 0 errors. Focused
  FileOps runtime archives
  `Specs\TestRuns\4cb089111a23\FileOps\2026-07-05_161924\` and
  `Specs\TestRuns\4cb089111a23\FileOps\2026-07-05_161955\` passed
  `Riptide_LiveFinishedSnapshotCarriesDiagnostics` and
  `Floodgate_CrossFsMoveCleanupDetectsDestinationCorruption`, respectively, each
  at 3 / 0 / 0.

- 2026-07-05 FolderView GR-P3 perf-budget tooling gate checkpoint:
  added `RunAllTestsPlan.Tests.ps1` coverage proving `-PerfBudgetPath` and
  `-RequirePerfBudgets` are threaded only into native self-test plan entries, plus
  `TestHarnessSourceContracts.Tests.ps1` coverage proving the native harness keeps
  FolderView perf budgets multi-machine and strict-mode gated. RED Pester runs
  first failed with the missing `PerfBudgetPath` parameter and missing
  `--selftest-require-perf-budgets` source contract; GREEN Pester passed
  `RunAllTestsPlan.Tests.ps1` at 10 / 0 / 0 and
  `TestHarnessSourceContracts.Tests.ps1` at 35 / 0 / 0. GREEN app build
  `.build\logs\msbuild-20260705_155942_502.log` produced
  `.build\x64\Debug\RedSalamander.exe`; GREEN runtime archive
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_160750\` exercised
  `Run-AllTests.ps1 -PerfBudgetPath Specs\Testing\FolderViewPerfBudgets.json5`
  and confirmed the multi-machine file was parsed before Debug skipped the
  Release-only hard budget.

- 2026-07-05 FolderView GR-24 draw-item brush-reuse guard checkpoint:
  made `folderView_draw_item_brush_reuse_guard` non-vacuous by keeping the normal
  selected/hovered render assertion at `drawItemTransientBrushCreateCount == 0` and
  adding a forced transient-brush positive control that must drive the same counter
  above zero. RED build `.build\logs\msbuild-20260705_153748_205.log` failed before
  `FolderWindow::DebugSetPaneForceDrawItemTransientBrushCreateForSelfTest(...)`
  existed; GREEN build `.build\logs\msbuild-20260705_153949_588.log` produced
  `.build\x64\Debug\RedSalamander.exe`. GREEN focused archive
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_154145\` passed
  `folderView_draw_item_brush_reuse_guard` with 1 / 0 / 0. GREEN adjacent archive
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_154806\` passed
  `folderView_perf_huge_folder_scale` with 1 / 0 / 0, preserving the zero-counter
  invariant at scale.

- 2026-07-05 FolderView GR-23 live IconCache failure retry checkpoint:
  added `folderView_iconcache_live_path_failure_retries_without_negative_cache`,
  covering a forced first live `QuerySysIconIndexForPath(...,
  useFileAttributes=false)` failure, no live-path negative-cache store, immediate
  shell re-entry on the second lookup, and positive-cache reuse on the third lookup.
  RED archive `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_151809\` first failed
  before the failure hook was wired; RED archive
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_152050\` then failed with
  `Second live icon lookup should re-query the shell and recover after a transient
  failure.` GREEN build `.build\logs\msbuild-20260705_152904_817.log` passed with
  0 warnings / 0 errors; GREEN focused archive
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_153136\` passed 1 / 0 / 0.
  Adjacent `folderView_perf_slow_virtual_provider` first failed at
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_152816\` because it still expected
  live failures to hit the old negative cache; after updating that guard, adjacent
  archives `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_153107\` and
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_153209\` passed the slow-provider
  and icon-pipeline guards.

- 2026-07-05 FolderView GR-22 refresh-to-paint stale-metric checkpoint:
  added `folderView_refresh_to_paint_metric_clears_after_failed_render`, covering a
  ready same-folder refresh-to-paint metric whose paint then fails at a synthetic
  non-device-loss `EndDraw` path, plus a later unrelated successful Present with no
  new refresh. RED archive
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_150226\` failed with
  `Stale refresh-to-paint metric emitted on an unrelated later Present; before=0
  after=1.`; GREEN build `.build\logs\msbuild-20260705_150338_519.log` passed with
  0 warnings / 0 errors; GREEN focused archive
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_150541\` passed 1 / 0 / 0.
  Adjacent rendering and refresh-metric archives
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_150654\`,
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_150739\`,
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_151107\`, and
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_151143\` passed the rendering
  error overlay, render device-loss recovery, refresh-preservation, and
  directory-change storm guards.
- 2026-07-05 Compare Directories GR-21 one-shot invert-selection checkpoint:
  extended `cmd_compare_directories_non_file_plugin_path_form_selection_and_empty_state`
  so a non-file plugin compare first observes the default left-only selection,
  invokes `IDM_COMPARE_INVERT_DIFFERENCES_SELECTION`, observes the inverted
  selection, then sends `IDM_LEFT_REFRESH` and requires the default compare
  selection to return. RED archive
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_143810\` failed with
  `leftSelected=0`; GREEN build `.build\logs\msbuild-20260705_144528_275.log`
  passed with 0 warnings / 0 errors; GREEN focused archive
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_144738\` passed 1 / 0 / 0;
  adjacent `cmd_compare_directories_` archive
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_145147\` passed 15 / 0 / 0.
- 2026-07-05 Settings hot-reload GR-20 async arming checkpoint:
  added `settings_hot_reload_transient_arm_failure_is_async`, covering one forced
  transient `FindFirstChangeNotificationW(...)` arming failure, a
  `SettingsHotReload::Start(...)` responsiveness budget under 200 ms, no
  `WAIT_TIMEOUT` hard-fail/teardown, watcher self-arming on retry, and a later
  settings change posting the settings-changed message. RED compile seam
  `.build\logs\msbuild-20260705_141406_457.log`; RED behavioral archive
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_141831\`; GREEN build
  `.build\logs\msbuild-20260705_142433_667.log`; GREEN focused archive
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_142640\`; GREEN adjacent
  `settings_hot_reload_` archive
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_142715\`.
- 2026-07-05 DxUi/FolderView GR-A6 shared device-loss checkpoint:
  added `TestDxUiDeviceLossAndD3dCreationAreShared`, covering the shared
  DxUi `IsDeviceLossHResult(...)` predicate, `DXGI_ERROR_DEVICE_HUNG`
  classification, WindowHost render-path routing through that predicate, and
  shared D3D11 hardware-to-WARP creation helper usage by WindowHost and
  FolderView. RED `.\build.ps1 -ProjectName DxUiTests -Configuration Debug
  -Platform x64` failed before implementation with missing
  `IsDeviceLossHResult`; GREEN DxUiTests build passed with 0 warnings / 0
  errors (`.build\logs\msbuild-20260705_140302_574.log`), focused
  `.build\x64\Debug\DxUiTests.exe --suite=WindowHost` exited 0, app build
  passed with 0 warnings / 0 errors
  (`.build\logs\msbuild-20260705_140513_048.log`), and focused Commands
  `folderView_render_device_loss_recovers` passed in
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_140722/`.

- 2026-07-05 FolderView GR-18/TW-4 thumbnail pending-accounting checkpoint:
  added `folderView_thumbnail_stale_bitmap_messages_preserve_pending_count`,
  covering stale batch/generation thumbnail bitmap payloads and late current
  bitmap deliveries that do not own a pending-count increment. RED focused
  rerun first failed at stale-message pending `0` vs expected `2`, then the
  TW-4 extension failed at pending `1` vs expected `2`; GREEN focused rerun
  passed after stale returns moved before the decrement and late/unaccounted
  payloads skipped pending decrement.

- 2026-07-05 DxUi GR-16 direct-root collapse parity checkpoint:
  added `TestAccessibilityLabelOnlyRootDoesNotUseDirectSemanticRootCollapse`,
  covering Label-only retained roots. RED
  `.\.build\x64\Debug\DxUiTests.exe --suite=Accessibility` failed at
  `label-only root exposes the label as a child instead of collapsing it into the root`.
  GREEN focused rerun passed after the snapshot builder recorded collapsed-root
  eligibility through the same single-semantic-root predicate as the retained path.

- 2026-07-05 DxUi GR-6 UIA dispatch timeout ownership checkpoint:
  added `TestAccessibilityUiActionDispatchOwnsTimedOutRequestStorage` and
  `TestAccessibilityTextRangeBoundingRectanglesTimeoutKeepsLateHandlerStorageAlive`,
  covering the cross-thread UIA action dispatch contract. The first source
  guard went RED because the shared heap dispatch helper was absent; the
  dynamic timeout case forces a dequeued `GetBoundingRectangles` handler to
  complete after the sender's 5s wait has returned `ERROR_TIMEOUT`, proving the
  caller output stays untouched while late handler output remains heap-owned.
  The folded GR-17 menu guard added
  `TestContextMenuDebugStateCrossThreadQueryDoesNotUseTimedOutStackStorage`,
  covering the test-only context-menu debug state path so it uses synchronous
  `SendMessageW(...)` instead of timed-out stack output storage.

- 2026-07-05 SearchService GR-S3 SQLite generation checkpoint:
  added `search_service_sqlite_external_rotation_refreshes_without_retry`,
  covering an externally rotated SQLite store during a running service session.
  RED archive
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_113058/`
  failed because the post-rotation query reached the `retry_query_ms` path.
  GREEN focused rerun
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_113650/`
  passed with 1 passed / 0 failed, proving query-time generation validation
  refreshes cached store info before the rotated query and avoids a
  post-rotation retry metric. Adjacent focused checks passed in
  `2026-07-05_114114/`, `2026-07-05_114131/`, and `2026-07-05_114218/`;
  live-journal-currentness guarded checks skipped in `2026-07-05_114043/`,
  `2026-07-05_114147/`, and `2026-07-05_114203/`.

- 2026-07-05 SearchService CW-10 transient-authorization cache checkpoint:
  added `search_service_transient_authorization_failure_is_incomplete_not_cached`,
  covering transient per-candidate directory authorization failures after an
  accepted service query. RED archive
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_111008/`
  failed with missing `ACCESS_DENIED_SKIPPED` (`warnings=0x00000000`).
  GREEN focused rerun
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_111440/`
  passed with 1 passed / 0 failed, proving the transient failure reports an
  incomplete-warning result and does not cache a false authorization decision
  for a later query in the same service session. Adjacent SearchService durable
  denied, candidate-authorization warning, progress, and status/query checks
  passed in `2026-07-05_111529/`, `2026-07-05_111546/`,
  `2026-07-05_111603/`, and `2026-07-05_111619/`.

- 2026-07-05 SearchService GR-14 candidate-authorization checkpoint:
  added `search_service_candidate_impersonation_failure_is_incomplete_warning`,
  covering a mid-stream service candidate authorization impersonation failure.
  RED archive
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_104801/`
  failed with `hr=0x80070558`. GREEN focused rerun passed with 1 passed / 0
  failed, proving the query completes with the unaffected candidate and
  `FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED`. Adjacent SearchService
  progress, cached-descendant authorization, status/query, multi-client
  rebuild, and deleted-root rebuild checks also passed. Archived runs:
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_105520/`,
  `2026-07-05_105549/`, `2026-07-05_105550/`, `2026-07-05_105551/`,
  `2026-07-05_105551_001/`, and `2026-07-05_105552/`.

- 2026-07-05 SearchService GR-11 deleted-root rebuild checkpoint:
  added `search_service_rebuild_deleted_root_purges_index`, covering the
  service rebuild/invalidate path after an indexed root directory is deleted.
  RED `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter search_service_rebuild_deleted_root_purges_index -FailFast -TimeoutMultiplier 3`
  failed with `SearchServiceBroker::RequestRebuild should purge a deleted root.
  hr=0x80070002`. GREEN focused rerun passed with 1 passed / 0 failed, and
  adjacent SearchService checks for existing-root rebuild, device-root
  rejection, status/query, and cached-descendant authorization also passed.
  Archived runs:
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_095517/`,
  `2026-07-05_095657/`, `2026-07-05_095746/`, `2026-07-05_095836/`, and
  `2026-07-05_095853/`.
- 2026-07-05 MTP GR-7 completed-swap journal checkpoint:
  Debug x64 `RedSalamander` build passed with
  `.build/logs/msbuild-20260705_092513_652.log` (`0 warning(s)`, `0
  error(s)`). RED
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_overwrite_journal_clears_completed_swap_without_temp -FailFast -TimeoutMultiplier 3`
  failed before the replay fix with `replay retained the stale
  completed-swap journal`. GREEN focused replay passed, the adjacent overwrite
  journal replay cases passed, and GREEN
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_ -FailFast -TimeoutMultiplier 3`
  passed with `50 passed / 0 failed / 1 skipped`. Added and listed
  `mtp_overwrite_journal_clears_completed_swap_without_temp`, which proves a
  retained no-tempPUID journal from a completed swap is cleared once the final
  destination size matches, and that the next read command avoids repeat replay
  sweep work.

- 2026-07-04 MTP GR-P1 WPD session/path-cache checkpoint:
  Debug x64 `RedSalamander` build passed with
  `.build/logs/msbuild-20260704_221619_628.log` (`0 warning(s)`, `0
  error(s)`). RED
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_wpd_session_and_path_cache_reuse -FailFast -TimeoutMultiplier 3`
  failed with `MTP WPD cache: create selftest instance failed.
  hr=0x8007007F`. GREEN
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_ -FailFast -TimeoutMultiplier 3`
  passed with `48 passed / 0 failed / 1 skipped`, archive
  `Specs/TestRuns/4cb089111a23/Mtp/2026-07-04_221500_gr_p1_wpd_session_path_cache/`.
  Added and listed `mtp_wpd_session_and_path_cache_reuse`, which proves the
  WPD backend reuses one device enumeration/session and hits the path cache for
  repeated metadata resolution.

- 2026-07-04 MTP GR-P1 worker-queue checkpoint:
  Debug x64 `RedSalamander` build passed with `0 warning(s), 0 error(s)`.
  RED
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_backend_command_worker_is_reused -FailFast -TimeoutMultiplier 3`
  failed with `expected one long-lived backend worker thread, observed 5`.
  GREEN
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_ -FailFast -TimeoutMultiplier 3`
  passed with `47 passed / 0 failed / 1 skipped`, archive
  `Specs/TestRuns/4cb089111a23/Mtp/2026-07-04_214700_gr_p1_worker_queue/`.
  Added and listed `mtp_backend_command_worker_is_reused`, which proves
  concurrent fake-backend metadata calls reuse one backend command worker while
  preserving single-transport serialization.

- 2026-06-28 Preferences/DxUi restore-clamp checkpoint:
  Debug x64 `RedSalamander` build passed with
  `.build/logs/msbuild-20260628_103428_253.log` (`0 warning(s)`, `0
  error(s)`). The focused Commands rerun for the Preferences Themes/DxUi
  blockers passed with `5 passed / 0 failed / 0 skipped`, archive
  `Specs/TestRuns/feb0d5542efb/Commands/2026-06-28_preferences_focus_after_restore_clamp/last_run/`.
  It covers bounded Themes long-list layout, pointer row selection, header drag
  and resize, and live DxUi interaction in the Plugins configuration page after
  clamping restored Preferences bounds to the effective minimum size. The
  follow-up full skip-build attempt is archived at
  `Specs/TestRuns/feb0d5542efb/FullSuite/2026-06-28_after_preferences_restore_clamp_full_skipbuild/last_run/`,
  but timed out before producing `run-all-tests-results.json`; partial artifacts
  show `DxUiTests` completed with all tests passed, `PerformanceTests2` passed
  12/12, and FileOps reached family 5 before the wrapper timeout. This is not
  full-suite closeout evidence.

- 2026-06-28 MTP review-fix follow-up checkpoint:
  Debug x64 `RedSalamander` build passed through the targeted plugin-deployment
  Pester child build with
  `.build/logs/msbuild-20260627_235143_567.log` (`0 warning(s), 0
  error(s)`). The rebuilt focused MTP Compare run
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_ -FailFast -TimeoutMultiplier 3`
  passed with `44 passed / 0 failed / 1 skipped`, archive
  `Specs/TestRuns/feb0d5542efb/CompareDirectories/2026-06-27_mtp_review_fixes_after_pathfix/last_run/`;
  the skip is the explicit no-device `mtp_live_device_smoke` skip. Follow-up
  path-length fixture fixes are focused-green for the two search-service
  ProgramData SQLite cases
  (`Specs/TestRuns/feb0d5542efb/FullSuite/2026-06-27_mtp_review_fixes_search_pd_focus/last_run/`,
  `2 passed / 0 failed`) and the Batch Rename parent/child cache case
  (`Specs/TestRuns/feb0d5542efb/FullSuite/2026-06-27_mtp_review_fixes_commands_brpc_focus/last_run/`,
  `1 passed / 0 failed`). At that checkpoint, full-suite closeout was still open: the archived full
  run under
  `Specs/TestRuns/feb0d5542efb/FullSuite/2026-06-27_mtp_review_fixes_full_after_mutex/last_run/`
  failed with `312 passed / 3 failed / 45 skipped`, and a later Commands-only
  run under
  `Specs/TestRuns/feb0d5542efb/FullSuite/2026-06-27_mtp_review_fixes_commands_full_after_pathfix/last_run/`
  still hit Preferences/DxUi failures before timing out.

- 2026-06-30 MTP deterministic closeout checkpoint:
  `.\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 3` passed with
  `1125 total / 1078 passed / 0 failed / 47 skipped`, archive
  `Specs/TestRuns/feb0d5542efb/Continuation/2026-06-30_223700_mtp_full_green_handoff/full_last_run/`.
  The corresponding Compare archive
  `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-30_214406/`
  passed with `205 passed / 0 failed / 25 skipped`; `mtp_live_device_smoke`
  is present and skipped explicitly because `REDSALAMANDER_SELFTEST_MTP_DEVICE`
  is not set. Its `perf/perf_metrics.jsonl` contains 647 `mtp.*` records
  across 40 unique metrics covering enum, props, transfer, writer, verify,
  path, overwrite, and device/watchdog families. The deterministic/accounting
  MTP perf gate is documented in
  `Specs/TestRuns/feb0d5542efb/Continuation/2026-06-30_223700_mtp_full_green_handoff/mtp_metric_audit.md`.
  A 2026-07-01 read-only production WPD discovery probe is archived at
  `Specs/TestRuns/feb0d5542efb/Continuation/2026-06-30_223700_mtp_full_green_handoff/live_probe_no_wpd_2026-07-01/`;
  it ran `mtp_live_device_smoke` with an impossible requested device name and
  skipped with `MTP live smoke skipped: no WPD/MTP devices enumerated.`
  `Tools/Run-MtpLiveCloseout.ps1` now wraps that case and archives command,
  stdout/stderr, env, PnP/CIM probes, and `last_run/`; its verified no-device
  probe archive is
  `Specs/TestRuns/feb0d5542efb/Continuation/2026-06-30_223700_mtp_full_green_handoff/live_probe_helper_no_wpd_2026-07-01/`
  with `total=1 / passed=0 / failed=0 / skipped=1`. After the 2026-07-01
  user-directed scope change, physical-device smoke, hotplug/device-removal
  behavior, WPD cancellation/stream-release proof, and live or recorded
  throughput evidence are optional post-closeout/manual validation, not MTP
  v1 closeout gates.

- 2026-06-27 MTP review-fix checkpoint:
  Debug x64 build passed with
  `.build/logs/msbuild-20260627_171217_892.log` (`0 warning(s), 0
  error(s)`). Focused
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_ -FailFast -TimeoutMultiplier 3`
  passed with `44 passed / 0 failed / 1 skipped`, archive
  `Specs/TestRuns/feb0d5542efb/CompareDirectories/2026-06-27_1718_mtp_review_fixes/focused_mtp_root/last_run/`.
  The checkpoint fixes and guards 8-byte `FileInfo` packing alignment,
  bounded WPD bulk-property callback waits, watchdog abandonment of wedged
  backend sessions, single-file `GetDirectorySize` metadata sizing, production
  shaped overwrite-temp orphan sweeps, real per-item batch completion indices,
  and picker preference for the device-object persistent id.

- 2026-06-27 MTP package harvest checkpoint:
  Release ZIP, MSIX, and MSI package harvests were verified for
  `Plugins\FileSystemMtp.dll` plus all four `Lang\FileSystemMtp-<culture>.dll`
  satellite DLLs. MSI evidence was captured by building
  `.build\AppPackages\RedSalamander-7.0.1-x64.msi` with WiX `6.0.2` and
  decompiling the package to
  `Specs/TestRuns/7d3a1247382a/PackageHarvest/2026-06-27_1323_mtp_package_harvest/RedSalamander-7.0.1-x64.decompiled.wxs`.
  `Tools\Tests\WingetValidation.Tests.ps1` now includes a directory-preserving
  ZIP/MSIX package harvest guard and passed with `17 passed / 0 failed / 0 skipped`.

- 2026-06-27 MTP watchdog cancel checkpoint:
  Debug x64 build passed with
  `.build/logs/msbuild-20260627_103043_961.log` (`0 warning(s), 0
  error(s)`). Focused
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_ -FailFast -TimeoutMultiplier 3`
  passed with `43 passed / 0 failed / 1 skipped`, archive
  `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-27_103528/`.
  Added and listed `mtp_watchdog_requests_backend_cancel`, which proves the
  MTP watchdog queues backend cancellation by using a fake backend whose
  delayed operation only unblocks when `RequestCancel()` is issued.

- 2026-06-27 MTP WPD write checkpoint:
  Debug x64 build passed with
  `.build/logs/msbuild-20260627_100045_446.log` (`0 warning(s), 0
  error(s)`). Focused
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_live_device_smoke -FailFast -TimeoutMultiplier 3`
  passed the wrapper with `0 passed / 0 failed / 1 skipped`, and focused
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_ -FailFast -TimeoutMultiplier 3`
  passed with `42 passed / 0 failed / 1 skipped`, archive
  `Specs/TestRuns/feb0d5542/CompareDirectories/2026-06-27_100701_mtp_wpd_writes_docs_refresh/`.
  `mtp_live_device_smoke` still skips before WPD by default, and its explicit
  opt-in path now owns scratch create/write/read/overwrite/rename/copy/move/delete
  coverage against the production WPD backend.

- 2026-07-05 MTP Connection Manager plugin-backed picker checkpoint:
  Debug x64 build passed with
  `.build/logs/msbuild-20260705_133324_146.log` (`0 warning(s), 0
  error(s)`). Focused
  `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_connection_manager_window_mtp_picker_populates_profile -FailFast -TimeoutMultiplier 3`
  passed with `1 passed / 0 failed / 0 skipped`, archive
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_133530/`. The archive's
  `perf/perf_metrics.jsonl` includes `mtp.connection_browse.devices`,
  `mtp.connection_browse.devices_us`, `mtp.connection_browse.storages`, and
  `mtp.connection_browse.storages_us`, proving the picker now routes through the
  MTP plugin browse export and the fake backend instead of host-side WPD picker
  fixtures.

- 2026-06-26 MTP Connection Manager picker checkpoint:
  Debug x64 build passed with
  `.build/logs/msbuild-20260626_181622_577.log` (`0 warning(s), 0
  error(s)`). Focused
  `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_connection_manager_window_mtp_picker_populates_profile -FailFast -TimeoutMultiplier 3`
  passed with `1 passed / 0 failed / 0 skipped`, archive
  `Specs/TestRuns/7613b115c/Commands/2026-06-26_162044_mtp_picker/`.
  Added and listed `cmd_connection_manager_window_mtp_picker_populates_profile`,
  which proves the MTP Connection Manager editor uses deterministic
  worker-backed picker data, hides auth/raw host/path fields, persists the WPD
  PnP hint plus `extra.devicePuid`/`friendlyName`/`readOnly`, and keeps the
  resource localization contract green for the new picker strings.

- 2026-06-26 MTP overwrite safety matrix checkpoint:
  Debug x64 build remained clean with
  `.build/logs/msbuild-20260626_165235_751.log` (`0 warning(s), 0
  error(s)`). The focused MTP Compare run
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_ -TimeoutMultiplier 3.0`
  passed with `42 passed / 0 failed / 1 skipped`, archive
  `Specs/TestRuns/7613b115c/CompareDirectories/2026-06-26_172010_mtp_overwrite_safety_matrix/`.
  Added and listed `mtp_overwrite_never_duplicates_or_halfwrites`, which
  aggregates the fake-backend overwrite crash-window matrix for temp upload
  failure, step-0 journal write failure, empty tempPUID policy block,
  rename-temp recovery, after-Commit orphan cleanup, and bounded
  unrenamable-temp terminal retention. The case asserts no duplicate final
  sibling, no half-written replacement, and only safe terminal states.

- 2026-06-26 MTP no-tempPUID journal sweep checkpoint:
  Debug x64 build remained clean with
  `.build/logs/msbuild-20260626_160614_636.log` (`0 warning(s), 0
  error(s)`). The focused MTP Compare run
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_ -TimeoutMultiplier 3.0`
  passed with `41 passed / 0 failed / 1 skipped`, archive
  `Specs/TestRuns/7613b115c/CompareDirectories/2026-06-26_161735_mtp_journal_no_temp_puid_sweep/`.
  Added and listed `mtp_overwrite_journal_recovers_committed_temp_without_tempPuid`,
  which covers retained overwrite-journal replay when the final exists, exact
  temp path is missing, and no `tempPuid` is available. The replay sweeps by
  random token, streamable metadata, declared size, and timestamp, deletes one
  exact candidate and clears the journal, retains ambiguous candidates with the
  journal, and emits `mtp.overwrite.journal_orphan_sweep_temp_removed` /
  `mtp.overwrite.journal_orphan_sweep_retained`.

- 2026-06-26 MTP journal temp cleanup retry checkpoint:
  Debug x64 build remained clean with
  `.build/logs/msbuild-20260626_144756_480.log` (`0 warning(s), 0
  error(s)`). The focused MTP Compare run
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_ -TimeoutMultiplier 3.0`
  passed with `40 passed / 0 failed / 1 skipped`, archive
  `Specs/TestRuns/7613b115c/CompareDirectories/2026-06-26_152300_mtp_journal_temp_cleanup_retry/`.
  Added and listed `mtp_overwrite_journal_replay_temp_cleanup_delete_failure_retries`,
  which covers retained overwrite-journal replay when both final and temp exist
  but temp cleanup deletion fails once, keeps the journal for retry, then
  removes the temp, preserves final bytes, clears the journal, and emits
  `mtp.overwrite.journal_replay_failed` followed by
  `mtp.overwrite.journal_replay_temp_removed`.

- 2026-06-26 MTP recovery rename rejection checkpoint:
  Debug x64 build remained clean with
  `.build/logs/msbuild-20260626_142033_480.log` (`0 warning(s), 0
  error(s)`). The focused MTP Compare run
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_ -TimeoutMultiplier 3.0`
  passed with `39 passed / 0 failed / 1 skipped`, archive
  `Specs/TestRuns/7613b115c/CompareDirectories/2026-06-26_142430_mtp_recovery_rename_rejection/`.
  Added and listed `mtp_overwrite_journal_replay_rename_rejection_is_bounded`,
  which covers bounded retained-journal retry when recovery rename is rejected,
  terminal journal clearing, no future infinite retry, retained temp readback,
  and `mtp.overwrite.journal_replay_retry_scheduled` /
  `mtp.overwrite.journal_replay_rename_terminal`.

- 2026-06-26 MTP temp-copy failure checkpoint:
  Debug x64 build remained clean with
  `.build/logs/msbuild-20260626_140713_480.log` (`0 warning(s), 0
  error(s)`). The focused MTP Compare run
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_ -TimeoutMultiplier 3.0`
  passed with `38 passed / 0 failed / 1 skipped`, archive
  `Specs/TestRuns/7613b115c/CompareDirectories/2026-06-26_141255_mtp_temp_copy_failure/`.
  Added and listed `mtp_copy_move_overwrite_temp_copy_failure_keeps_original_and_allows_retry`,
  which covers copy and move overwrite temp-copy failures after journal intent,
  preserves source/destination bytes, leaks no temp sibling, clears the
  journal, allows later retry, and emits `mtp.overwrite.temp_copy_failed`.

- 2026-06-26 MTP temp-upload failure checkpoint:
  Debug x64 build remained clean with
  `.build/logs/msbuild-20260626_135526_639.log` (`0 warning(s), 0
  error(s)`). The focused MTP Compare run
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_ -TimeoutMultiplier 3.0`
  passed with `37 passed / 0 failed / 1 skipped`, archive
  `Specs/TestRuns/7613b115c/CompareDirectories/2026-06-26_140330_mtp_temp_upload_failure/`.
  Added and listed `mtp_overwrite_temp_upload_failure_keeps_original_and_allows_retry`,
  which covers a writer overwrite temp-upload failure after journal intent,
  preserves original bytes, leaks no temp sibling, clears the journal, allows a
  later retry, and emits `mtp.overwrite.temp_upload_failed`.

- 2026-06-26 MTP journal replay orphan-temp checkpoint:
  Debug x64 build remained clean with
  `.build/logs/msbuild-20260626_111702_204.log` (`0 warning(s), 0
  error(s)`). The focused MTP Compare run
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_ -TimeoutMultiplier 3.0`
  passed with `36 passed / 0 failed / 1 skipped`, archive
  `Specs/TestRuns/7613b115c/CompareDirectories/2026-06-26_135009_mtp_journal_replay_orphan_temp/`.
  Added and listed `mtp_overwrite_journal_replay_removes_temp_when_final_exists`,
  which covers retained overwrite-journal replay when both final and temp
  exist, removes the random temp sibling, preserves final bytes, clears the
  host journal, and emits `mtp.overwrite.journal_replay_temp_removed`.

- 2026-06-26 MTP journal replay rename-temp checkpoint:
  Debug x64 build remained clean with
  `.build/logs/msbuild-20260626_105205_768.log` (`0 warning(s), 0
  error(s)`). The focused MTP Compare run
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_`
  passed with `35 passed / 0 failed / 1 skipped`, archive
  `Specs/TestRuns/7613b115c/CompareDirectories/2026-06-26_105806_mtp_journal_replay_rename_temp/`.
  Added and listed `mtp_overwrite_journal_recovers_rename_temp_failure`,
  which covers retained overwrite-journal replay after a one-shot
  temp-to-final rename failure, verifies exactly one final entry and no leaked
  `.rs-mtp-overwrite-` sibling, and emits
  `mtp.overwrite.rename_temp_failed`,
  `mtp.overwrite.journal_replay_temp_committed`, and
  `mtp.overwrite.journal_cleared`.

- 2026-06-26 MTP overwrite-journal preflight checkpoint:
  Debug x64 build remained clean with
  `.build/logs/msbuild-20260626_102610_090.log` (`0 warning(s), 0
  error(s)`). The focused MTP Compare run
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_`
  passed with `34 passed / 0 failed / 1 skipped`, archive
  `Specs/TestRuns/7613b115c/CompareDirectories/2026-06-26_103012_mtp_journal_preflight/`.
  Added and listed `mtp_overwrite_journal_write_failure_aborts_before_upload`,
  which covers writer, copy, and move overwrite journal-write preflight
  failures failing closed before backend write/copy/move, preserving
  source/destination contents, leaking no temp sibling, and emitting
  `mtp.overwrite.journal_write_failed`. The same checkpoint fixed the
  CompareDirectories harness exit crash by releasing and reacquiring
  manager-owned plugin COM interfaces around MTP runtime-refresh coverage.

- 2026-06-26 MTP copy/move overwrite temp-swap checkpoint:
  Debug x64 build remained clean with
  `.build/logs/msbuild-20260625_184027_401.log` (`0 warning(s), 0
  error(s)`). The focused MTP Compare run
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_`
  passed with `33 passed / 0 failed / 1 skipped`, archive
  `Specs/TestRuns/7613b115c/CompareDirectories/2026-06-26_093936/`.
  Added and listed `mtp_copy_move_overwrite_uses_temp_puid_swap`, which
  covers copy and move overwrite replacement bytes, final PUID handoff, source
  retention/deletion semantics, no leaked temp siblings, and
  `mtp.overwrite.device_source_temp_swap_committed`.

- 2026-06-28 FolderView icon-pipeline recall avoidance and metric coverage:
  Debug `RedSalamander` build passed with zero warnings/errors
  (`.build/logs/msbuild-20260628_192440_115.log`). The new
  `folderView_perf_icon_pipeline_cold_slow` Commands case first failed at
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_190646/` because
  offline/recall per-file icon lookup still consumed live path lookups. After
  the fix, Debug passed at
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_192653/`, with
  `iconPathLiveLookupConsumeCount=0`, `recallAvoidedMetricRows=3`, and
  `thumbnailProviderAllowedConsumeCount=0`. Supporting Debug FolderView cases
  passed: `folderView_perf_iconcache_contention`
  (`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_192709/`),
  `folderView_perf_cold_first_visit`
  (`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_192722/`), and
  `folderView_perf_slow_virtual_provider`
  (`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_192733/`). The
  PerformanceTests2 icon enumeration tests passed through `vstest.console.exe`
  (`LargeFolderIconEnumeration_DuplicatePaths` and
  `LargeFolderIconEnumeration_MixedItems`). Test-enabled Release build passed
  with zero errors and the existing optimized-build C4883 warning in
  `FolderWindow.FileOperations.SelfTest.cpp`
  (`.build/logs/msbuild-20260628_192752_991.log`). Release
  `folderView_perf_icon_pipeline_cold_slow` passed and emitted
  `build=Release` at
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_193018/`; Release support
  cases passed at `2026-06-28_193033/`, `2026-06-28_193045/`, and
  `2026-06-28_193058/`.

- 2026-06-28 FolderView perf budget gates: test-enabled Release
  `RedSalamander` build passed with the existing optimized-build warning C4883
  in File Operations selftests and zero errors
  (`.build/logs/msbuild-20260628_181130_414.log`). The budgeted six-case
  Commands run with `--selftest-perf-budget=Specs\Testing\FolderViewPerfBudgets.json5`
  passed (`6 passed / 0 failed`), archive
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_181518/`. A scratch
  impossible-budget smoke check failed as expected at
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_180631/` with
  `FolderView perf budget folder.frame.total_us.p95 exceeded`, proving the hard
  budget path fails the case instead of only logging. Debug `RedSalamander`
  build passed after the same changes with zero warnings/errors
  (`.build/logs/msbuild-20260628_181748_075.log`).

- 2026-06-21 CompareDirectories Crosscut closeout audit:
  `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2`
  completed in 66m 7.1s and failed overall (`991 passed / 13 failed / 50
  skipped`), so
  `Specs/Plans/WIP/Operation_Crosscut_CompareDirectoriesRemediation_SyncDataSafetyAndOptionsSimplification_2026-06-15.md`
  remains in WIP. The CompareDirectories leg passed (`140 passed / 0 failed /
  24 skipped`), archive
  `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-21_002337/`.
  Commands failed (`750 passed / 10 failed / 2 skipped`), archive
  `Specs/TestRuns/7d3a1247382a/Commands/2026-06-21_004849/`, with failures
  in Preview-pane tab close-button visibility, Batch Rename reporting/cancel
  cases, Find history hover repaint, FileOps conflict diagnostics,
  hidden/system visibility toggle stability, and persistent menu hover
  switching. FileOps failed (`90 passed / 2 failed / 24 skipped`), archive
  `Specs/TestRuns/7d3a1247382a/FileOps/2026-06-21_012000/`, in
  `Phase5_PreCalcCancelReleasesSlot` and
  `Phase11_BridgeMultiFolderParallelCopyInFlightLines`. DxUiTests exited 1 in
  `TestWindowHostEmitsFrameStageMetricsForCaptureRender`; details are in
  `%LOCALAPPDATA%\RedSalamander\SelfTest\last_run\DxUiTests.output.log`.

- 2026-06-20 CompareDirectories CX4 standing-test closeout:
  Debug `RedSalamander` build passed with
  `.build/logs/msbuild-20260620_233426_995.log` (`0 warning(s), 0
  error(s)`). Added and listed six standing Compare cases:
  `content_equal_size_equal_mtime_differs`,
  `content_unknown_size_streaming_compare`, `content_cache_hit_skips_io`,
  `setRoots_resets_and_bumps_version`, `select_subdirs_only_in_one_pane`, and
  `missing_side_empty_enumeration`. Focused runs passed in archives
  `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-20_233323/`,
  `2026-06-20_233335/`, `2026-06-20_233803/`,
  `2026-06-20_233813/`, `2026-06-20_233821/`, and
  `2026-06-20_233826/`. The `content_cache_hit_skips_io` case initially
  failed at
  `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-20_233340/`,
  exposing that `SetSettings` cleared completed `_contentCompareCache` entries;
  the fix preserves completed content-cache entries for settings-only
  recomputes while root changes and invalidation still clear/evict them. Full
  `Run-AllTests.ps1 -Suite Compare -SkipBuild -TimeoutMultiplier 2` passed
  (`164 total / 140 passed / 0 failed / 24 skipped`), archive
  `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-20_234056/`.
  Focused command-dispatch proof passed for `cmd_compare_directories_`
  (`15 passed / 0 failed`), archive
  `Specs/TestRuns/7d3a1247382a/Commands/2026-06-21_001914/`, and
  `cmd_preferences_dialog_compare_directories_page_uses_dxui_statics`
  (`1 passed / 0 failed`), archive
  `Specs/TestRuns/7d3a1247382a/Commands/2026-06-21_001936/`.

- 2026-06-20 Compare window controller-state extraction (CX3-4): Debug
  `RedSalamander` build passed with
  `.build/logs/msbuild-20260620_231739_525.log` (`0 warning(s), 0
  error(s)`). Inspection confirmed chrome/menu/banner state is grouped under
  `_chrome`, Options dialog/body/draft/theme/scroll state under
  `_optionsPanel`, and progress banner/ETA/spinner/watermark/task-card state
  under `_progress`; the old flat `_compareRunSawScanProgress`,
  `_bannerRescanIsCancel`, `_compareTaskId`, `_compareRunResultHr`,
  `_watermarkState`, `_options*`, `_dxMenuBar*`, `_dxBanner*`, and
  `_scanProgress*` state families are gone. The Win32 callback friend shims
  remain as message-dispatch adapters. Focused
  `cmd_compare_directories_window_uses_dxui_menu_bar_and_banner_buttons`
  passed (`1 passed`), archive
  `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_232056/`. Full
  `cmd_compare_directories_options_` Commands family passed (`10 passed / 0
  failed`), archive
  `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_232125/`. Focused
  `cmd_compare_directories_progress_perf` passed (`1 passed`), archive
  `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_232133/`. Full
  `Run-AllTests.ps1 -Suite Compare -SkipBuild -TimeoutMultiplier 2` passed
  (`158 total / 134 passed / 0 failed / 24 skipped`), archive
  `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-20_232352/`.
  A first parallel validation attempt hit the documented self-test mutex with
  exit 3 for overlapping jobs; the serial reruns above are the recorded green
  evidence.

- 2026-06-20 Compare Options draft-only state model (CX3-3): Debug
  `RedSalamander` build passed with
  `.build/logs/msbuild-20260620_230155_024.log` (`0 warning(s), 0
  error(s)`). Inspection confirmed the Options body no longer creates hidden
  native statics, checkboxes, or edits as a state model; `_optionsUi` owns only
  the scroll/container host, live DxUi toggles/edits write directly to the
  `Common::Settings::CompareDirectoriesSettings` draft, and OK persists that
  draft while Cancel discards it. Grep confirmed removal of
  `OptionsToggleCard`, `OptionsIgnoreCard`, `SetTwoStateToggleState`,
  `GetTwoStateToggleState`, `syntheticHiddenToggle`, `_optionsFrameStyle`, and
  `ThemedInputFrames`, plus no hidden Button/Edit creation in
  `CompareDirectoriesWindow.Options.cpp`. Same-process DxUi UIA ValuePattern
  reads/writes now use message-pumped helpers so the UI thread services retained
  control providers while tests inspect values. Full
  `Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter
  cmd_compare_directories_options_ -TimeoutMultiplier 2` passed (`10 passed /
  0 failed`), archive
  `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_231049/`. Full
  `Run-AllTests.ps1 -Suite Compare -SkipBuild -TimeoutMultiplier 2` passed
  (`158 total / 134 passed / 0 failed / 24 skipped`), archive
  `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-20_230833/`.

- 2026-06-20 Compare Options split-host compatibility cleanup (CX3-2):
  Debug `RedSalamander` build passed with
  `.build/logs/msbuild-20260620_222034_595.log` (`0 warning(s), 0
  error(s)`). Inspection confirmed the never-attached
  `OptionsSectionDxLabel`, `OptionsCardDxText`, `OptionsToggleDx`, and
  `OptionsEditDx` compatibility structs/members are gone, and the fallback
  layout path is legacy-HWND-only after the live body-host branch. Focused
  `Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter
  cmd_compare_directories_options_uses_dxui_labels_without_visible_legacy_statics
  -TimeoutMultiplier 2` passed (`1 passed`), archive
  `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_222945/`. Focused
  `cmd_compare_directories_options_scroll_to_lower_cards_stays_stable`
  passed (`1 passed`), archive
  `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_222955/`. Focused
  `cmd_compare_directories_options_enter_and_escape_route_default_cancel`
  passed (`1 passed`), archive
  `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_223010/`. Full
  `Run-AllTests.ps1 -Suite Compare -SkipBuild -TimeoutMultiplier 2` passed
  (`158 total / 134 passed / 0 failed / 24 skipped`), archive
  `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-20_223229/`.
  A broad `cmd_compare_directories_options_` prefix attempt exceeded a 300s
  tool timeout and is not counted as green evidence.

- 2026-06-20 Compare menu/options dedup and test-only decision API (CX3-5):
  Debug `RedSalamander` build passed with
  `.build/logs/msbuild-20260620_221106_949.log` (`0 warning(s), 0
  error(s)`). Inspection confirmed Compare Directories menu conversion now uses
  `DxUiNativeMenuInterop`, the duplicate options footer button-host attach is
  gone, and `CompareDirectoriesSession::GetOrComputeDecision()` is available
  only under `ENABLE_TESTS`. Focused `Run-AllTests.ps1 -Suite Commands
  -SkipBuild -CaseFilter
  cmd_compare_directories_window_uses_dxui_menu_bar_and_banner_buttons
  -TimeoutMultiplier 2` passed (`1 passed`), archive
  `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_221500/`. Focused
  `Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter
  cmd_compare_directories_options_uses_dxui_labels_without_visible_legacy_statics
  -TimeoutMultiplier 2` passed (`1 passed`), archive
  `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_221517/`. Full
  `Run-AllTests.ps1 -Suite Compare -SkipBuild -TimeoutMultiplier 2` passed
  (`158 total / 134 passed / 0 failed / 24 skipped`), archive
  `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-20_221733/`.

- 2026-06-20 Compare content unsupported option persistence (CX2-9):
  Debug `RedSalamander` build passed with
  `.build/logs/msbuild-20260620_220311_534.log` (`0 warning(s), 0
  error(s)`). Inspection confirmed unsupported content compare clears the
  disabled Options toggle and `ReadOptionsControlsToSettings()` masks
  persisted `compareContent` with `_session->IsContentCompareSupported()`.
  Focused `Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter
  content_no_io_disables_compareContent -TimeoutMultiplier 2` passed (`1
  passed`), archive
  `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-20_220654/`.
  Focused `Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter
  cmd_compare_directories_non_file_plugin_path_form_selection_and_empty_state
  -TimeoutMultiplier 2` passed (`1 passed`), archive
  `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_220703/`.

- 2026-06-20 Compare Options wheel remainder reset (CX2-8): Debug
  `RedSalamander` build passed with
  `.build/logs/msbuild-20260620_215541_640.log` (`0 warning(s), 0
  error(s)`). Inspection confirmed non-scrollable Win32/DxUi wheel paths
  and options layout recomputation reset `_optionsWheelRemainder`.
  Focused `Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter
  cmd_compare_directories_options_scroll_to_lower_cards_stays_stable
  -TimeoutMultiplier 2` passed (`1 passed`), archive
  `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_215932/`.

- 2026-06-20 Compare task-card trailing flush (CX2-7): Debug
  `RedSalamander` build passed with
  `.build/logs/msbuild-20260620_215041_527.log` (`0 warning(s), 0
  error(s)`). Inspection confirmed throttled task-card updates schedule a
  one-shot trailing flush and cancel it on applied update, finish, dismiss,
  and teardown. Focused `Run-AllTests.ps1 -Suite Commands -SkipBuild
  -CaseFilter cmd_compare_directories_progress_perf -TimeoutMultiplier 2`
  passed (`1 passed`), archive
  `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_215432/`.

- 2026-06-20 Compare ETA sanity clamp (CX2-6): Debug
  `RedSalamander` build passed with
  `.build/logs/msbuild-20260620_214522_317.log` (`0 warning(s), 0
  error(s)`). Inspection confirmed content ETA display now requires a
  finite rate of at least 1 KiB/s and a finite ETA no larger than 7 days.
  Focused `Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter
  cmd_compare_directories_progress_perf -TimeoutMultiplier 2` passed (`1
  passed`), archive
  `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_214910/`.

- 2026-06-20 Compare decision-refresh timer fallback (CX2-5): Debug
  `RedSalamander` build passed with
  `.build/logs/msbuild-20260620_213922_473.log` (`0 warning(s), 0
  error(s)`). Inspection confirmed `SetTimer` failure logs
  `Debug::ErrorWithLastError(...)` and posts
  `WndMsg::kCompareDirectoriesDecisionRefreshNow` with duplicate fallback
  suppression. Focused normal-path `Run-AllTests.ps1 -Suite Commands
  -SkipBuild -CaseFilter cmd_compare_directories_progress_perf
  -TimeoutMultiplier 2` passed (`1 passed`), archive
  `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_214325/`.

- 2026-06-20 Compare animated watermark pulse independence (CX2-4):
  Debug `RedSalamander` build passed with
  `.build/logs/msbuild-20260620_213339_896.log` (`0 warning(s), 0
  error(s)`). Focused `Run-AllTests.ps1 -Suite Commands -SkipBuild
  -CaseFilter cmd_compare_directories_progress_perf -TimeoutMultiplier 2`
  passed (`1 passed`), archive
  `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_213739/`.

- 2026-06-20 Compare leave-scope empty-previous root fallback (CX2-3):
  code inspection confirmed `SyncOtherPanePath(...)` falls back through
  `previousPath.value_or(rootPath)` and revalidates the fallback before
  posting the deferred prompt. Focused `Run-AllTests.ps1 -Suite Commands
  -SkipBuild -CaseFilter
  cmd_compare_directories_leave_scope_prompt_defers_out_of_navigation_callback
  -TimeoutMultiplier 2` passed (`1 passed`), archive
  `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_212725/`.

- 2026-06-20 Compare ignored-directory direct-navigation guard (CX2-2):
  Debug `RedSalamander` build passed with
  `.build/logs/msbuild-20260620_211808_480.log` (`0 warning(s), 0
  error(s)`). Focused `Run-AllTests.ps1 -Suite Compare -SkipBuild
  -CaseFilter ignore_direct_navigation_subtree` passed (`1 passed`),
  archive
  `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-20_212156/`.
  Full Compare passed (`158 total / 134 passed / 0 failed / 24 skipped`),
  archive
  `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-20_212418/`.

- 2026-06-20 Compare normalized-name collision guard (CX2-1): Debug
  `RedSalamander` build passed with
  `.build/logs/msbuild-20260620_210758_239.log` (`0 warning(s), 0
  error(s)`). Focused `Run-AllTests.ps1 -Suite Compare -SkipBuild
  -CaseFilter normalized_name_collision_preserves_same_side_entries`
  passed (`1 passed`), archive
  `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-20_211136/`.
  Full Compare passed (`157 total / 133 passed / 0 failed / 24 skipped`),
  archive
  `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-20_211403/`.

- 2026-06-20 Compare create-failure ownership guard (CX1-3): Debug
  `RedSalamander` build passed with
  `.build/logs/msbuild-20260620_202843_064.log` (`0 warning(s), 0
  error(s)`). Focused `Run-AllTests.ps1 -Suite Commands -SkipBuild
  -CaseFilter cmd_compare_directories_create_failure_does_not_double_delete
  -TimeoutMultiplier 2` passed (`1 passed`), archive
  `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_203227/`.

- 2026-06-20 Compare leave-scope prompt reentrancy guard (CX1-2): Debug
  `RedSalamander` build passed with
  `.build/logs/msbuild-20260620_201727_700.log` (`0 warning(s), 0
  error(s)`). Focused `Run-AllTests.ps1 -Suite Commands -SkipBuild
  -CaseFilter cmd_compare_directories_leave_scope_prompt_defers_out_of_navigation_callback
  -TimeoutMultiplier 2` passed (`1 passed`), archive
  `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_202105/`.

- 2026-06-20 Compare decision-cache pending-budget guard (CX1-1):
  Debug `RedSalamander` build passed with
  `.build/logs/msbuild-20260620_194901_078.log` (`0 warning(s), 0
  error(s)`). Focused `Run-AllTests.ps1 -Suite Compare -SkipBuild
  -CaseFilter decision_cache_eviction_budget_pending_wide_tree
  -TimeoutMultiplier 2` passed (`1 passed`), archive
  `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-20_195237/`.
  The archived perf JSONL includes
  `compare.selftest.decision_cache.pending_budget_bytes` with
  `budget=32768 allowed=131072`, current bytes `31350`, and high-water
  bytes `34596`. Full Compare passed (`156 total / 132 passed / 0 failed /
  24 skipped`), archive
  `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-20_195612/`.
- 2026-06-20 Compare sync manifest closeout (CX0-1): Debug
  `RedSalamander` build passed with
  `.build/logs/msbuild-20260620_191715_299.log` (`0 warning(s), 0
  error(s)`). Focused `Run-AllTests.ps1 -Suite Compare -SkipBuild
  -CaseFilter sync_manifest_ -TimeoutMultiplier 2` passed (`4 passed`),
  archive
  `Specs/TestRuns/LT-PF5VDAGE/Compare/2026-06-20_192106_sync_manifest_cx0_1_focus_final/`.
  Full Compare passed (`155 total / 131 passed / 0 failed / 24 skipped`),
  archive
  `Specs/TestRuns/LT-PF5VDAGE/Compare/2026-06-20_192323_compare_full_cx0_1_final/`.
  FileOps Clearflow Phase 4 passed (`6 passed`), archive
  `Specs/TestRuns/LT-PF5VDAGE/FileOps/2026-06-20_192518_resolved_items_route_cx0_1_final/`.
  The archived perf JSONL contains `compare.sync.manifest.*` build,
  item, blocker, elapsed, and submitted-item counters.
- 2026-06-09 FolderView DPI/scale repaint guard: Debug `RedSalamander`
  build passed with `.build/logs/msbuild-20260609_133818_682.log` (`0
  warning(s), 0 error(s)`). `.build/x64/Debug/RedSalamander.exe
  --commands-selftest --selftest-case=folderView_dpi_change_repaints_both_panes`
  exited 0; archive
  `Specs/TestRuns/7d3a1247382a/Commands/2026-06-09_134243/`. The case
  covers dual-pane DPI propagation, swap-chain resize recovery, and the
  required full-client repaint boundary when moving between different monitor
  scales.
- 2026-06-09 FolderView rendering-error overlay persistence: Debug
  `RedSalamander` build passed with
  `.build/logs/msbuild-20260609_133818_682.log` (`0 warning(s), 0 error(s)`).
  `.build/x64/Debug/RedSalamander.exe --commands-selftest
  --selftest-case=folderView_rendering_error_overlay_requires_persistence`
  exited 0; archive
  `Specs/TestRuns/7d3a1247382a/Commands/2026-06-09_134254/`. The case
  covers transient rendering-failure suppression, persistent rendering-alert
  promotion, explicit clear/reset behavior, and the
  `folder.render.failure_suppressed` / `folder.render.failure_promoted`
  metric pair.
- 2026-06-09 DxUi menu popup live-DPI relayout: Debug `DxUiTests`
  build passed with `.build/logs/msbuild-20260609_113137_699.log`
  (`0 warning(s), 0 error(s)`). `.build/x64/Debug/DxUiTests.exe
  --suite=Menu` exited 0 after adding
  `TestContextMenuPopupRelayoutsOnDpiChanged`, covering an open async context
  menu that receives `WM_DPICHANGED`, recomputes its DPI-scaled visible
  surface/window capture geometry, keeps all visible rows inside the viewport,
  and closes cleanly through Escape. The same pass also keeps the stationary
  mouse keyboard-root-switch guard green after captured-popup stale moves stop
  seeding the owner root-switch point. Perf evidence is archived at
  `Specs/TestRuns/4cb089111a23/DxUiTests/2026-06-09_113318_menu_dpi_popup_relayout/`,
  including `dxui.menu.dpi_relayout_us`.
- 2026-06-08 DxUi button/menu chrome and AlertOverlay closeout: Debug
  `DxUiTests` build passed with
  `.build/logs/msbuild-20260608_172213_729.log` (`0 warning(s), 0
  error(s)`), and `.build/x64/Debug/DxUiTests.exe --suite=NewControls`
  exited 0 with the shared split/dropdown layout and custom AlertOverlay
  style guards green. Debug `RedSalamander` build passed with
  `.build/logs/msbuild-20260608_172231_835.log` (`0 warning(s), 0
  error(s)`). Focused command selftests archived green at
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-08_172503/`
  (`cmd_app_prompt_uses_alert_overlay_window`, `1 passed`),
  `.../2026-06-08_172521/`
  (`cmd_pane_find_dialog_result_shortcuts_use_shell_clipboard_and_file_actions`,
  `1 passed`), `.../2026-06-08_172527/`
  (`cmd_compare_directories_window_uses_dxui_menu_bar_and_banner_buttons`,
  `1 passed`), `.../2026-06-08_172540/`
  (`cmd_plugin_configuration_dialog_uses_dxui_`, `2 passed`), and
  `.../2026-06-08_172546/`
  (`cmd_preferences_dialog_themes_page_uses_dxui_shell_chrome`,
  `1 passed`). The operation-window popup smoke rerun is archived at
  `Specs/TestRuns/4cb089111a23/FileOps/2026-06-08_172552/` (`3 passed`,
  `0 failed`). The broad native-control audit bundle is archived at
  `Specs/TestRuns/4cb089111a23/Audit/2026-06-08_172148/`, with
  `Audit-RemainingWin32UiDependencies.ps1 -FailOnFindings`,
  `Audit-VisibleNativeSurfaces.ps1`, and `Audit-ComctlReportSurfaces.ps1`
  all exiting 0. Release `RedSalamander` build passed with
  `.build/logs/msbuild-20260608_172738_771.log` (`0 warning(s), 0
  error(s)`), covering the warning gate for this change.
- 2026-06-08 ViewerSpace synthetic-bucket snapshot metric fix: Debug
  `RedSalamander` build passed with
  `.build/logs/msbuild-20260608_114330_167.log` (`0 warning(s), 0
  error(s)`). Red evidence archived at
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-08_114210/`: the new
  `cmd_viewer_space_synthetic_bucket_metrics_match_root_bytes` guard failed
  because retained scan "Other" plus transient layout "Other" reported
  `syntheticCount=2` for a files-only root that should have exactly one
  retained bucket. The focused green run archived at
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-08_114530/` (`1 passed`).
  The filtered `cmd_viewer_space_` sweep archived at
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-08_114653/` (`15 passed`,
  `1 skipped` opt-in), including the synthetic-byte reconciliation guard.
- 2026-06-07 ViewerSpace renderer resize/device-loss polish: Debug
  `RedSalamander` build passed with
  `.build/logs/msbuild-20260607_150528_070.log` (`0 warning(s), 0
  error(s)`). Focused command runs used visible `Start-Process -Wait` with
  `--commands-selftest --selftest-case=...`: forced device-loss recovery
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-07_150722/` (`1 passed`)
  forces both `D2DERR_RECREATE_TARGET` and `DXGI_ERROR_DEVICE_REMOVED`
  through the paint path and waits for a recreated device-context renderer.
  The final `cmd_viewer_space_` sweep archived at
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-07_150827/` (`14 passed`,
  `1 skipped` opt-in) includes the resize fast-path guard proving
  `rendererBrushCreateCount` and `rendererTextFormatCreateCount` stay
  unchanged while `swapChainResizeCount` advances.
- 2026-06-07 ViewerSpace TurboTreemap review follow-up: Debug
  `RedSalamander` build passed with
  `.build/logs/msbuild-20260607_132437_315.log` (`0 warning(s), 0
  error(s)`). Focused red evidence included the renderer resize guard at
  `Specs/TestRuns/4cb089111a23/SelfTest/2026-06-07_131727/`, where the
  client grew to `1184x780` but the swap chain stayed `978x644`, and the
  animated hit-grid guard at
  `Specs/TestRuns/4cb089111a23/SelfTest/2026-06-07_132323/`, where one
  sampled point differed from the reverse-linear fallback. The corrected
  focused matrix passed with visible `Start-Process -Wait` runs:
  renderer resize `2026-06-07_132749`, adaptive scan budget
  `2026-06-07_132751`, dense-file LOD/detail `2026-06-07_132752`, animated
  hit-grid parity `2026-06-07_132754`, serial/parallel byte golden
  `2026-06-07_132801`, queue storm `2026-06-07_132807`, hover static-cache
  stability `2026-06-07_132809`, and the source lifecycle guard
  `2026-06-07_132810`.
- 2026-06-06 ViewerSpace large-sibling detail follow-up: Debug
  `RedSalamander` build passed with
  `.build/logs/msbuild-20260606_235706_934.log` (`0 warning(s), 0
  error(s)`). Filtered runner inventory now reports 13 `cmd_viewer_space_`
  cases and 2 `viewer_space_perf_` cases, including
  `cmd_viewer_space_large_sibling_folders_expose_nested_detail`. The red
  archive at `Specs/TestRuns/4cb089111a23/Commands/2026-06-06_235254/`
  showed the 14 large sibling folders staying flat with
  `viewer.space.layout.visible_tiles=14` and
  `viewer.space.render.tile_draw_count=14` after scanning 896 files. The
  post-change run exited `0` at
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-06_235859/` and exposed
  nested detail with `viewer.space.layout.visible_tiles=910`,
  `viewer.space.render.tile_draw_count=910`,
  `viewer.space.layout.rebuild_us=2365`, `viewer.space.render.paint_us=5811`,
  `viewer.space.hit_grid.max_candidates=25`, and
  `viewer.space.queue.pending_bytes=0`.
- 2026-06-06 ViewerSpace TurboTreemap closeout: Debug `RedSalamander`
  build passed with `.build/logs/msbuild-20260606_205744_268.log` (`0
  warning(s), 0 error(s)`). Filtered runner inventory reports 12
  `cmd_viewer_space_` cases and 2 `viewer_space_perf_` cases. The full
  `cmd_viewer_space_` sweep exited `0` at
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-06_210644/` (`12 passed, 0
  failed, 0 skipped`), including device-context readiness, static-cache hover,
  dense-file retention, spatial-grid parity, serial-vs-parallel scanner totals,
  cancellation/teardown stress, scan-cache cap/skip, queue storm, memory
  settle, and opt-in 20k visible-tile coverage. Its perf rows include
  `viewer.space.model.file_candidate_count=20000`,
  `viewer.space.render.tile_draw_count=20000`,
  `viewer.space.layout.visible_tiles=20000`,
  `viewer.space.hit_grid.cells=276`,
  `viewer.space.hit_grid.max_candidates=90`,
  `viewer.space.queue.pending_bytes=0`,
  `viewer.space.queue.coalesced_count=47`,
  `viewer.space.model.cache_skipped_large=83366`, and
  `viewer.space.scan.active_workers=8`. The `viewer_space_perf_` sweep exited
  `0` at `Specs/TestRuns/4cb089111a23/Commands/2026-06-06_210710/`; the
  bounded default small-progressive scenario passed and the large scenario
  skipped by design without `REDSALAMANDER_VIEWERSPACE_LARGE_PERF=1`.
- 2026-06-06 ViewerSpace TurboTreemap implementation slice: Debug
  `RedSalamander` build passed with
  `.build/logs/msbuild-20260606_191021_398.log` (`0 warning(s), 0
  error(s)`). `--selftest-list-cases --commands-selftest
  --selftest-case=cmd_viewer_space_` now lists five focused ViewerSpace command
  cases, including
  `cmd_viewer_space_renderer_device_context_ready_resize_close`,
  `cmd_viewer_space_hover_does_not_rebuild_static_cache`, and
  `cmd_viewer_space_dense_files_exposes_more_than_legacy_topk`. Focused runs
  exited `0` and archived under
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-06_191215/`,
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-06_191225/`, and
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-06_191228/`. The post-change
  `viewer_space_perf_small_progressive` run exited `0` at
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-06_191303/`; its scenario
  artifact reports renderer mode `DeviceContext`, `10000` retained file
  candidates, `20` visible tiles, `0` culled tiles, `last_paint_us=1081`,
  `last_layout_us=80`, and `last_working_set_bytes=147644416`.
- 2026-06-06 ViewerSpace Phase 0 perf instrumentation: Debug `RedSalamander`
  build passed with `.build/logs/msbuild-20260606_182424_290.log`
  (`0 warning(s), 0 error(s)`). The rebuilt runner lists two Commands perf
  cases for the `viewer_space_perf_` prefix:
  `viewer_space_perf_small_progressive` and `viewer_space_perf_large_optin`.
  The focused archived run for `viewer_space_perf_small_progressive` exited
  `0` and archived `commands_results.json`, `commands_trace.txt`,
  `perf/perf_metrics.jsonl`, and
  `perf/viewer_space_perf_small_progressive_metrics.json` under
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-06_182818/`.
- 2026-06-05 Find Look-in NavigationView/result context menu coverage:
  Debug `RedSalamander` build passed with
  `.build/logs/msbuild-20260605_191128_989.log` (`0 warning(s), 0 error(s)`),
  Debug `DxUiTests` build passed with
  `.build/logs/msbuild-20260605_184731_385.log` (`0 warning(s), 0 error(s)`),
  and `.build\x64\Debug\DxUiTests.exe --suite=Grid` exited 0. The focused
  command case
  `cmd_pane_find_dialog_result_shortcuts_use_shell_clipboard_and_file_actions`
  passed at `Specs/TestRuns/4cb089111a23/Commands/2026-06-05_191333/`.
  This run covers the embedded `Look in` `NavigationView`, the hidden root
  combo remaining non-visual plumbing, result-grid right-click preservation of
  multi-selection, the two-section result context menu, shortcut text resolved
  from active `ShortcutManager` bindings, clicked-item dispatch operating only
  on the hit row, and selection-section dispatch operating on the whole selected
  result set. The broad `DxUiTests --suite=Menu` sweep was not recorded as green
  in this desktop run: repeated attempts failed in pre-existing timing-sensitive
  menu cases before reaching the Find result-menu assertions.
- 2026-06-04 central pointer input router continuation closeout: the
  `cmd_viewer_` shutdown blocker is closed. ViewerSpace now exposes a module
  quiet point and process-exit retention path for DLL-global scheduler/cache,
  window-class, and graphics state; `cmd_viewer_` passed on the current Debug
  app build at `Specs/TestRuns/7d3a1247382a/Commands/2026-06-04_215054/`
  (`4 passed, 0 failed`). Debug `RedSalamander` build passed with
  `.build/logs/msbuild-20260604_213301_479.log` (`0 warning(s), 0 error(s)`).
  Debug `DxUiTests` rebuild passed with
  `.build/logs/msbuild-20260604_214754_056.log` (`0 warning(s), 0 error(s)`),
  and `.build\x64\Debug\DxUiTests.exe --suite=Menu` passed twice after the
  menu DPI clamp test was hardened to anchor outside the virtual screen instead
  of assuming no monitor exists left/above the primary monitor. Focused command
  slices passed at `2026-06-04_145308/` (`cmd_app_menuBar_`, `17 passed`),
  `2026-06-04_145534/` (`cmd_pane_navigation_`, `31 passed`),
  `2026-06-04_145538/` (pane status-bar sort popup), `2026-06-04_145551/`
  (file-operations speed-limit prompt, `4 passed`), and `2026-06-04_145554/`
  (current-directory context menu). The full Find prefix passed at
  `Specs/TestRuns/7d3a1247382a/Commands/2026-06-04_214327/` with `56 passed`,
  `0 failed`, and `6 skipped` because the desktop OS clipboard was unavailable
  (`OpenClipboard error=5, openWindow=0x0, owner=0x0`). The skipped cases are
  the six clipboard-content assertions only; the broader Find status,
  incremental-search, destination routing, action-button, layout, sort, and
  activation cases executed and passed. Earlier same-day clipboard coverage for
  the result shortcut case passed before the external clipboard lock at
  `Specs/TestRuns/7d3a1247382a/Commands/2026-06-04_180358/`, and the focused
  unavailable-clipboard run is archived at `2026-06-04_212254/`. The final
  source guard `powershell -NoProfile -ExecutionPolicy Bypass -File
  .\Tools\Tests\VerifyNoProductionGetCursorPos.Tests.ps1` reported
  `No production GetCursorPos violations found.`; `git diff --check -- Common
  RedSalamander Plugins Specs\UI Specs\Testing Scripts Tests` exited 0 with
  only line-ending normalization warnings.
- 2026-06-03 central pointer input router final closeout coverage
  (supersedes earlier same-day live-cursor classification notes for the final
  contract): production routing under `Common`, `RedSalamander`, and `Plugins`
  now treats delivered pointer message coordinates or explicit owner/control
  anchors as authoritative. The whole-tree source guard
  `powershell -NoProfile -ExecutionPolicy Bypass -File .\Tools\Tests\VerifyNoProductionGetCursorPos.Tests.ps1`
  first failed with 30 production `GetCursorPos()` violations, then passed with
  `No production GetCursorPos violations found.` after the migration. Raw
  `GetCursorPos()` remains only in selftests or same-line annotated
  diagnostic-only production trace sites. Debug `DxUiTests` build passed with
  `.build/logs/msbuild-20260603_201812_256.log` (`0 warning(s), 0 error(s)`),
  and `.build\x64\Debug\DxUiTests.exe --suite=Menu` exited 0 across three
  consecutive reruns plus final verification after fixing the Menu-suite flakes
  by settling delivered hover input, accepting either pending or already-fired
  delayed submenu timer state, and parking cursor state for keyboard/baseline
  probes. Debug `RedSalamander` build passed with
  `.build/logs/msbuild-20260603_201952_657.log` (`0 warning(s), 0 error(s)`).
  Focused command selftests exited 0 after one setup-only cursor-position retry:
  the retry pass is archived at
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-03_202320/`, and the final
  green batch runs from `2026-06-03_202334/` through
  `2026-06-03_202346_001/` for the Find destination, Find popup/Escape,
  result-shortcut/help overlay, and ViewerText/ViewerSpace delivered
  anchor/hover cases. NavigationView now routes delivered owner pointer events
  from target HWND and client-point metadata only; stale edit hosts and owner
  residue are rejected by target/capture/message-order and explicit teardown
  checks instead of a synthetic input-generation token.
  Final operator-style live validation used the Debug app with
  `REDSALAMANDER_DXUI_MENU_TRACE=1` and per-case
  `REDSALAMANDER_DXUI_MENU_TRACE_FILE` outputs under
  `Specs/TestRuns/4cb089111a23/LivePointer/2026-06-03_final_validation/`.
  Thirteen command archives, `2026-06-03_203537/`, `2026-06-03_203539/`,
  `2026-06-03_203542/`, `2026-06-03_203548/`, `2026-06-03_203549/`,
  `2026-06-03_203550/`, `2026-06-03_203633/`, `2026-06-03_203634/`,
  `2026-06-03_203635/`, `2026-06-03_203636/`, `2026-06-03_203637/`,
  `2026-06-03_203639/`, and `2026-06-03_203640/`, all exited 0. That live slice
  covers the Find split-button menu opening, hover highlight, outside
  light-dismiss, destination NavigationView history/menu routing with same-owner
  pointer drift, stale edit-host retirement, help overlay close glyph, Escape,
  backdrop repaint, FolderWindow current-directory context routing, pane
  status-bar delivered hover/sort-click popup open/close, file-operation
  speed-limit popup/prompt ownership and close, and ViewerText/ViewerSpace
  context-menu and hover routing from delivered points rather than later cursor
  state.
- 2026-06-03 central pointer input router: focused baseline command selftests
  passed before migration (`cmd_pane_find_dialog_destination_navigation_stale_edit_host_hit_testing`,
  `cmd_pane_find_dialog_result_shortcuts_use_shell_clipboard_and_file_actions`,
  `cmd_pane_find_dialog_result_drains_respect_child_input_queue_order`,
  `cmd_pane_find_dialog_escape_closes_popup_before_cancel`, and
  `cmd_pane_find_dialog_escape_from_dx_control_closes_cancel`, all `EXIT=0`).
  The new whole-tree guard
  `powershell -NoProfile -ExecutionPolicy Bypass -File .\Tools\Tests\VerifyNoProductionGetCursorPos.Tests.ps1`
  is intentionally red at the start of the migration, reporting 30 production
  `GetCursorPos()` violations across `Common`, `RedSalamander`, and `Plugins`.
- 2026-06-03 Find destination near-owner queued click classification:
  the live repro log showed `WM_MOUSEMOVE` / `WM_LBUTTONDOWN` delivered to the
  destination `NavigationView` with an inside-nav client point, while
  `GetCursorPos()` had already drifted a few pixels below the control on the
  same Find owner window. Treating that as stale cleared hover and dropped the
  click, so menus appeared to wait for unrelated pointer movement. The
  contract now distinguishes same-owner edge drift from true stale residue:
  same-owner fringe messages are accepted, while far or different-window live
  pointers are still ignored/cleared. The new red probe in
  `cmd_pane_find_dialog_destination_navigation_stale_edit_host_hit_testing`
  failed at `Specs/TestRuns/4cb089111a23/Commands/2026-06-03_153437/` with
  "Find destination navigation should accept an inside-history queued click
  when the live cursor only drifted to the same-owner fringe." Debug
  `RedSalamander` build passed with
  `.build/logs/msbuild-20260603_153841_661.log` (`0 warning(s), 0 error(s)`).
  Focused green coverage is archived at
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-03_154056/`.
- 2026-06-03 Find destination stale queued activation and active-edit
  chrome routing: live repro logs showed a queued destination `NavigationView`
  `WM_LBUTTONDBLCLK` whose delivered client point was still over the footer
  path area, but whose live cursor had already left the child HWND by the time
  the message was processed. Accepting that stale double-click entered
  destination edit mode and left the edit host owning later pointer messages
  until unrelated title-bar or owner-window movement changed focus. The final
  contract is that no-capture `NavigationView` owner hover/click/double-click
  messages are ignored and hover is cleared when `GetCursorPos()` /
  `WindowFromPoint()` prove the live pointer is far from the control or on a
  different top-level window; small same-owner edge drift remains valid
  delivered input. Stale double-clicks must not enter edit mode. Active edit mode
  still keeps menu/history/disk buttons live, and inactive stale edit hosts
  retire as transparent/hidden before destination history/menu input is routed.
  The red command run at
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-03_140743/` caught the stale
  double-click entering edit mode. Debug `RedSalamander` build passed with
  `.build/logs/msbuild-20260603_150035_996.log` (`0 warning(s), 0 error(s)`).
  Focused green coverage is archived at
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-03_150245/`. Adjacent
  queue/popup/Escape regressions also passed at
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-03_145844/`,
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-03_145844_001/`, and
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-03_145845/`.
- 2026-06-03 DxUi menu delivered popup root-switch routing: the refreshed live
  repro showed popup/captured mouse messages whose delivered `lParam` no longer
  matched `GetCursorPos()` by the time the menu loop handled them. The new red
  `TestMenuPopupMouseMoveUsesDeliveredPointForRootSwitch` failed because
  `RouteMenuPointerHover(...)` replaced the delivered popup `WM_MOUSEMOVE`
  point with the live cursor before probing top-level root switching. After
  removing that live-cursor substitution, popup and owner mouse messages use the
  delivered coordinate as the authoritative input point; the existing
  menu-bar-hover notification remains the explicit captured-menu-bar switching
  path. The same update drained generated cursor-move residue in synthetic idle
  tests so `TestMenuRootSwitchDoesNotPollCursorWhileIdle` validates no-message
  behavior instead of depending on the removed substitution. Debug `DxUiTests`
  build passed with `.build/logs/msbuild-20260603_094341_655.log`
  (`0 warning(s), 0 error(s)`), and `.build\x64\Debug\DxUiTests.exe --suite=Menu`
  exited 0. Debug `RedSalamander` build passed with
  `.build/logs/msbuild-20260603_094600_595.log` (`0 warning(s), 0 error(s)`),
  and the focused command selftests archived passed runs at
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-03_094812/`,
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-03_094825/`,
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-03_094837/`, and
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-03_094851/`.
- 2026-06-03 Find destination delivered pointer routing
  (superseded later on 2026-06-03 by stale queued activation handling above): a follow-up live repro
  log showed `WM_MOUSEMOVE` and `WM_LBUTTONDOWN` reaching the embedded
  destination `NavigationView`, with each delivered message `lParam` inside the
  child client area, but `NavigationView` rejected them as
  `live-cursor-outside` because `GetCursorPos()` had already moved outside the
  36px footer bar by handler time. The corrected red
  `cmd_pane_find_dialog_result_shortcuts_use_shell_clipboard_and_file_actions`
  assertion failed at
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-03_085637/` with
  `history=0`. After restoring Win32 delivered-message semantics so
  `WM_MOUSEMOVE` / `WM_LBUTTONDOWN` use their `lParam` client point and
  `GetCursorPos()` is limited to explicit timer/menu-loop reconciliation, Debug
  `RedSalamander` build passed with
  `.build/logs/msbuild-20260603_085733_299.log` (`0 warning(s), 0 error(s)`)
  and the same case passed at
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-03_085932/`. A later follow-up
  extended the destination hover probe so an immediate repaint after delivered
  `WM_MOUSEMOVE` must preserve the delivered hover even when the live cursor is
  outside the embedded footer; `WM_PAINT` and `WM_SETCURSOR` no longer mutate
  NavigationView hover state.
- 2026-06-03 Find destination stale edit-host hit-testing: live repro logs showed
  only `dxui.windowhost.raw` records for a 30px `RedSalamander.NavigationView`
  edit host covering the embedded destination bar, with no `navigation.wndproc`
  input reaching the history/menu branch. The new red
  `cmd_pane_find_dialog_destination_navigation_stale_edit_host_hit_testing`
  case recreated a stale visible `RedSalamander.NavigationView.DxHost` child
  and failed at
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-03_075606/` because the real
  mouse click on the destination history arrow was swallowed. A follow-up
  no-diagnostics run reproduced the same case as timing-sensitive at
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-03_103028/`: the stale edit host
  could receive activation and disappear before a forwarded click reached the
  history branch. After making inactive edit hosts hide during layout and
  retire/forward pointer activation from the triggering `WM_MOUSEACTIVATE`
  coordinates when edit focus has escaped, Debug `RedSalamander` builds passed with
  `.build/logs/msbuild-20260603_080013_833.log` and
  `.build/logs/msbuild-20260603_103349_929.log`; the final cleanup build passed
  with `.build/logs/msbuild-20260603_103856_705.log` (`0 warning(s), 0 error(s)`),
  the focused case passed at
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-03_080150/`, and related
  NavigationView history-keyboard and Find queue-order cases passed at
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-03_080208/` and
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-03_080231/`. The hardened
  focused case passed once at
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-03_103550/` and then three
  consecutive no-diagnostics reruns at
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-03_103603/`,
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-03_103605/`, and
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-03_103606/`; the final
  post-cleanup focused run passed at
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-03_104057/`. The Debug
  `DxUiTests --suite=Menu` suite passed with
  `.build/logs/msbuild-20260603_103618_258.log`.
- 2026-06-02 Find destination stale-hover/stale-click stabilization
  (superseded on 2026-06-03 by delivered pointer routing above): live repro
  logs showed a stale `WM_MOUSEMOVE` clearing correctly, followed by a stale
  `WM_LBUTTONDOWN` over the embedded destination disk/history area while the
  live cursor was already outside the `NavigationView`; that delayed click
  opened a menu only after unrelated pointer movement. The red
  `cmd_pane_find_dialog_result_shortcuts_use_shell_clipboard_and_file_actions`
  stale-click probe failed with "Find destination navigation should ignore
  stale click messages whose old coordinate no longer matches the live cursor."
  That fix shared the live-pointer resolver across hover and click activation,
  but was later found to be the wrong contract for delivered child-window
  messages; the 2026-06-03 entry above supersedes it.
  The same case also keeps the deterministic overlay close/backdrop coverage:
  owner/anchor backdrop capture is available before first paint, and close-glyph
  mouse messages dismiss without title-bar movement. Verification: Debug
  `RedSalamander` build passed with
  `.build/logs/msbuild-20260602_220733_842.log` (`0 warning(s), 0 error(s)`),
  the focused Find command selftest passed and archived
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-02_220930/`, Debug
  `DxUiTests` build passed with `.build/logs/msbuild-20260602_221207_344.log`
  (`0 warning(s), 0 error(s)`), and `.build\x64\Debug\DxUiTests.exe --suite=Menu`
  exited 0.
- 2026-06-02 Find destination stale-hover stabilization
  (superseded on 2026-06-03 by delivered pointer routing above): live repro logs showed
  queued `WM_MOUSEMOVE` coordinates over the embedded destination history arrow
  while the live cursor was outside the `NavigationView`. The new red
  `cmd_pane_find_dialog_result_shortcuts_use_shell_clipboard_and_file_actions`
  probe failed with `history=1` at
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-02_210659/`. That run changed
  `NavigationView::OnMouseMove` to compute hover from the live cursor/window hit
  result, but the delivered-message rule in the 2026-06-03 entry above now
  supersedes that behavior for `WM_MOUSEMOVE` messages routed to the child HWND.
  Debug `RedSalamander` build passed with
  `.build/logs/msbuild-20260602_210935_990.log` (`0 warning(s), 0 error(s)`),
  the same command selftest exited 0 and archived
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-02_211132/`, Debug
  `DxUiTests` build passed with `.build/logs/msbuild-20260602_211247_680.log`
  (`0 warning(s), 0 error(s)`), and `.build\x64\Debug\DxUiTests.exe --suite=Menu`
  exited 0. Earlier no-op destination path/history refresh guards remain covered
  by the same Find command case.
- 2026-06-02 NavigationView/Find destination menu diagnostics: Debug `RedSalamander` build passed with `.build/logs/msbuild-20260602_193607_073.log` (`0 warning(s), 0 error(s)`), Debug `DxUiTests` rebuild passed with `.build/logs/msbuild-20260602_193857_319.log` (`0 warning(s), 0 error(s)`), and `.build\x64\Debug\DxUiTests.exe --suite=Menu` exited 0 after the diagnostic-only input-state/edit-host/hover-timer/render tracing refresh.
- 2026-06-02 Find/DxUi raw input diagnostics: after the repro log showed no click/dropdown messages, Debug `RedSalamander` build passed with `.build/logs/msbuild-20260602_200651_211.log` (`0 warning(s), 0 error(s)`), Debug `DxUiTests` rebuild passed with `.build/logs/msbuild-20260602_200842_610.log` (`0 warning(s), 0 error(s)`), and `.build\x64\Debug\DxUiTests.exe --suite=Menu` exited 0 after adding diagnostic-only `find.wndproc.raw` and `dxui.windowhost.*` trace records.
- 2026-06-02 Find/DxUi raw input diagnostic flood trim: after live repro logging showed 368,340 `WM_NCHITTEST` and 38,900 owner `WM_MOUSELEAVE` trace records, Debug `RedSalamander` build passed with `.build/logs/msbuild-20260602_202216_918.log` (`0 warning(s), 0 error(s)`), Debug `DxUiTests` rebuild passed with `.build/logs/msbuild-20260602_202352_999.log` (`0 warning(s), 0 error(s)`), and `.build\x64\Debug\DxUiTests.exe --suite=Menu` exited 0 after removing reentrant/high-volume hit-test and leave logging.
- `DxUiTests.Grid` covers clipped-only fallback grid text tooltips, repeated explicit cell tooltips that reappear when the visible text is clipped, and the shared folder-view grid visual mode used by Find results.
- `cmd_pane_find_dialog_result_shortcuts_use_shell_clipboard_and_file_actions`
  now covers the embedded `Look in` NavigationView, two-section Find result
  context menu, clicked-item and whole-selection menu dispatch, effective
  shortcut text resolved from `ShortcutManager`, the right-aligned Find result
  `?` help button, modal help focus,
  synchronous close-glyph hit geometry before paint, visible first-frame modal
  help paint without mouse movement, captured owner backdrop for the modal
  scrim so the owner area cannot render black, nonzero applied scrim opacity,
  captured-backdrop refresh after owner resize while the help overlay is
  visible, delivered close-glyph press/release dismissal without title-bar movement,
  release-only close-glyph dismissal for first-activation routing, `Escape`
  dismissal, blank idle action-row status, blended action-row status text, bottom destination `NavigationView`
  visibility/text/embedded presentation/history seeding/delivered-pointer
  history-arrow popup behavior, rendered first-frame destination popup state,
  footer-top anchored compact scroll-capped embedded destination dropdowns,
  explicit F5/F6 destination override, destination text in
  the help message, no scheduled Find foreground/focus restoration after
  configured viewer/editor action launches, UI Automation context cleanup, and
  selftest shutdown cleanup for icon/D2D resources used by Find result icons.
- `cmd_pane_clipboardPaste_uses_preferred_move_effect` covers the real
  `cmd/pane/clipboardPaste` dispatch using `CF_HDROP` plus
  `Preferred DropEffect = DROPEFFECT_MOVE`, proving Ctrl+X then Ctrl+V moves
  source files instead of copying them and asks for exactly one shared move
  confirmation. It also waits for verified completion to invalidate the
  unchanged move clipboard and proves a second paste cannot retry the move.
- `folderView_drop_integrity_guards` covers hovered-subfolder versus pane-background
  destination resolution, no-modifier same-volume MOVE selection, external queued
  MOVE reporting COPY while the host request remains MOVE, same-parent/descendant
  rejection, and rejection while the destination folder is not enumerated.
- `folderView_drop_rejects_malformed_payloads` injects hostile private-drop and
  `CF_HDROP` HGLOBAL layouts and proves they fail before queuing a file operation.
- `folderView_pointer_targets_and_stale_hover_are_safe` covers plain-click selection
  collapse, empty-background drag disarm, and paint with an out-of-range stale hover.
- `TestHarnessSourceContracts.Tests.ps1` pins the remaining IronLedger seams that are
  impractical to allocate or interleave in a deterministic runtime fixture: bounded
  clipboard writer/reader arithmetic, untrusted provider count clamping, named focus
  preservation, nested-menu target revalidation, inline deferred dispatch, local-source
  provider routing, completion callback ownership, and the settled cross-pane gate.
- 2026-07-16 IronLedger/Observatory Track 4 closeout evidence is archived at
  `Specs/TestRuns/4cb089111a23/FolderView/2026-07-16_2127_observatory_track4_ironledger/`.
  The focused Debug build completed with 0 warnings/0 errors; the five selected Commands
  regressions exited 0; source contracts passed 134/134.
- 2026-06-28 FolderView refresh preservation metrics: Debug `RedSalamander`
  build passed with `.build/logs/msbuild-20260628_184739_282.log` (`0 warning(s),
  0 error(s)`). `folderView_perf_directory_change_storm` passed in
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_184947/`, covering broad
  create/rename/delete churn and archiving all seven `folder.refresh.*` rows.
  `folderView_perf_refresh_preservation` passed in
  `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_185004/`, covering one create,
  one delete, one same-folder rename, a bounded burst, focus preservation,
  active incremental-search preservation, unchanged selection preservation, and
  selected-rename transfer; its artifact records one row each for preserve,
  rebuild, selection-preserve, rename-transfer, request-to-paint, debounce, and
  enumeration metrics.

Related documents:
- `Specs/Testing/Testing_SelfTests.md` — result contract
- `Specs/Testing/Testing_PerformanceValidation.md` — perf validation requirements
- `Specs/Testing/Testing_SelfTestRemoteCredentials.md` — remote credential setup
- `Tests/README.md` — test infrastructure index

Current test-review evidence:

- Frame performance closeout, Task 12 (2026-05-19, machine `4cb089111a23`,
  branch `codex/dxui-frame-performance`) used per-command logs under
  `Specs/TestRuns/local_scratch/frame_perf_closeout_20260519_165526/`.
  Focused Debug validation:
  - `.\build.ps1 -Configuration Debug` exited 0; build log
    `.build/logs/msbuild-20260519_165527_622.log`; diagnostics 0 warnings,
    0 errors.
  - `.\.build\x64\Debug\DxUiTests.exe` exited 0; log
    `Specs/TestRuns/local_scratch/frame_perf_closeout_20260519_165526/02_debug_dxuitests.log`;
    output ended with `All DxUi tests passed`.
  - `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_large_folder_baseline --selftest-timeout-multiplier=4`
    exited 0; archive
    `Specs/TestRuns/4cb089111a23/Commands/2026-05-19_165927/`.
  - `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4`
    exited 0; archive
    `Specs/TestRuns/4cb089111a23/Commands/2026-05-19_165933/`.
  - `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_sort_toggle_stress --selftest-timeout-multiplier=4`
    exited 0; archive
    `Specs/TestRuns/4cb089111a23/Commands/2026-05-19_165945/`.
  - `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_iconcache_contention --selftest-timeout-multiplier=4`
    exited 0; archive
    `Specs/TestRuns/4cb089111a23/Commands/2026-05-19_165950/`.
  - `.\.build\x64\Debug\MonitorTest.exe` exited 0; log
    `Specs/TestRuns/local_scratch/frame_perf_closeout_20260519_165526/07_debug_monitortest.log`;
    no repo archive is expected for this entrypoint.
  - `try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Debug } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }`
    exited 0; build log `.build/logs/msbuild-20260519_165815_580.log`;
    diagnostics 0 warnings, 0 errors.
  - `.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --perf`
    exited 0 when run as a waited visible GUI process; archive
    `Specs/TestRuns/4cb089111a23/Monitor/2026-05-19_170035/` passed and
    contained required monitor frame metrics. A hidden-window wait is not a
    valid chrome selftest mode; it failed visibility/render checks at
    `Specs/TestRuns/4cb089111a23/Monitor/2026-05-19_165950/`.
- Frame performance Release evidence from the same closeout:
  - `.\build.ps1 -ProjectName RedSalamander -Configuration Release` exited 0;
    build log `.build/logs/msbuild-20260519_170103_776.log`; diagnostics
    0 warnings, 0 errors.
  - The exact normal Release command
    `.\.build\x64\Release\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4`
    exited 2 with no archive because normal Release `RedSalamander.exe` omits
    `ENABLE_TESTS`.
  - An explicit test-enabled Release rebuild for perf evidence,
    `try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamander -Configuration Release } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }`,
    exited 0; build log `.build/logs/msbuild-20260519_170412_660.log`;
    diagnostics 1 warning, 0 errors (`C4883` in
    `FolderWindow.FileOperations.SelfTest.cpp`).
  - After that test-enabled Release rebuild,
    `.\.build\x64\Release\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4`
    exited 0; archive
    `Specs/TestRuns/4cb089111a23/Commands/2026-05-19_170631/`.
  - `try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Release } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }`
    exited 0; build log `.build/logs/msbuild-20260519_170241_261.log`;
    diagnostics 0 warnings, 0 errors.
  - `.\.build\x64\Release\RedSalamanderMonitor.exe --chrome-selftest --perf`
    exited 1; archive
    `Specs/TestRuns/4cb089111a23/Monitor/2026-05-19_170251/` failed with
    `Monitor DxUI chrome selftest requires ENABLE_TESTS.` This was the red
    evidence for the Release Monitor test-definition wiring fix.
  - Follow-up fix `fix(monitor): enable release selftest opt-in` updated
    Release x64/ARM64 Monitor `ClCompile` definitions to include
    `$(RSBuildTestDefinitions)`. The rerun
    `try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Release } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }`
    exited 0; build log `.build/logs/msbuild-20260519_171244_915.log`;
    wrapper log
    `Specs/TestRuns/local_scratch/frame_perf_release_monitor_fix_20260519_171500/01_release_monitor_test_enabled_build.log`;
    diagnostics 0 warnings, 0 errors.
  - After that fix,
    `.\.build\x64\Release\RedSalamanderMonitor.exe --chrome-selftest --perf`
    exited 0; wrapper log
    `Specs/TestRuns/local_scratch/frame_perf_release_monitor_fix_20260519_171500/02_release_monitor_chrome_perf.log`;
    archive `Specs/TestRuns/4cb089111a23/Monitor/2026-05-19_171308/`
    passed with `monitorFrameMetricPresence.allPresent=true`.
  - Controller rerun evidence after the same fix:
    `try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Release } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }`
    exited 0; build log `.build/logs/msbuild-20260519_171503_045.log`;
    diagnostics 0 warnings, 0 errors. The rerun
    `.\.build\x64\Release\RedSalamanderMonitor.exe --chrome-selftest --perf`
    exited 0; archive `Specs/TestRuns/4cb089111a23/Monitor/2026-05-19_171520/`
    passed with `monitorFrameMetricPresence.allPresent=true`.
- Frame-performance ARM64 Release build validation from 2026-05-20:
  - `.\build.ps1 -ProjectName DxUiTests -Configuration Release -Platform ARM64`
    exited 0; build log `.build/logs/msbuild-20260520_211504_073.log`;
    diagnostics 0 warnings, 0 errors.
  - `.\build.ps1 -ProjectName RedSalamander -Configuration Release -Platform ARM64`
    exited 0; build log `.build/logs/msbuild-20260520_211647_017.log`;
    diagnostics 26 warnings, 0 errors. Warnings were existing Release ARM64
    padding warnings plus Monitor C5245 test-helper warnings when the monitor is
    compiled without `ENABLE_TESTS`.
  - `try { $env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Release -Platform ARM64 } finally { Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue }`
    exited 0; build log `.build/logs/msbuild-20260520_211828_344.log`;
    diagnostics 0 warnings, 0 errors.
  - ARM64 executable selftests were not run locally because this machine is
    `PROCESSOR_ARCHITECTURE=AMD64` on an AMD Ryzen 9 9950X3D
    (`Win32_Processor.Architecture=9`), which does not provide native ARM64
    process execution. The ARM64 evidence for this pass is build-only.
- Frame-performance remaining-work closeout from 2026-05-20:
  - After restoring the repo's x64 vcpkg triplet state,
    `.\build.ps1 -Configuration Debug` exited 0; build log
    `.build/logs/msbuild-20260520_212049_103.log`; diagnostics
    0 warnings, 0 errors. The restore was needed because the prior ARM64
    validation left the install root without x64 headers such as `wil/com.h`.
  - `.\build.ps1 -ProjectName DxUiTests -Configuration Debug` exited 0; build
    log `.build/logs/msbuild-20260520_212944_364.log`; diagnostics
    0 warnings, 0 errors. The full `.\.build\x64\Debug\DxUiTests.exe`
    run exited 0 and printed `All DxUi tests passed.`
  - `folderView_perf_overlay_invalidation_stress` exited 0; archive
    `Specs/TestRuns/4cb089111a23/Commands/2026-05-20_213228/`;
    `commands_results.json` reports 1 passed, 0 failed, 0 skipped.
  - `folderView_perf_scroll_render_stress` exited 0; archive
    `Specs/TestRuns/4cb089111a23/Commands/2026-05-20_213241/`;
    `commands_results.json` reports 1 passed, 0 failed, 0 skipped.
  - Test-enabled `RedSalamanderMonitor` Debug build exited 0; build log
    `.build/logs/msbuild-20260520_213246_343.log`; diagnostics
    0 warnings, 0 errors. Default chrome selftest archive
    `Specs/TestRuns/4cb089111a23/Monitor/2026-05-20_213256/` passed with
    `monitorFrameMetricPresence.allPresent=true` and
    `monitorScrollbackSelfTest.enabled=false`.
  - Monitor ETW latency archive
    `Specs/TestRuns/4cb089111a23/Monitor/2026-05-20_213301/` passed with
    `monitorEtwBurstLatency.metricPresence.allPresent=true`; p95 rows were
    `append_to_visible=22,501us`, `batch_drain=1,214us`,
    `frame.total=18,746us`, and `present=4,622us`.
  - Monitor scrollback archive
    `Specs/TestRuns/4cb089111a23/Monitor/2026-05-20_213308/` passed with
    `monitorScrollbackSelfTest.metricPresence.allPresent=true`; p95 rows were
    `scrollback_slice=35,844us`, `frame.total=128,975us`, and
    `present=867us`.
  - Cleanup included hiding Monitor selftest-only helpers behind
    `ENABLE_TESTS`, stabilizing the overlapping-popup menu test so the full
    DxUi suite is not order-sensitive to OS cursor movement, and removing the
    stale disposable baseline worktree
    `C:\Users\eric\.config\superpowers\worktrees\RedSalamander\dxui-overlay-baseline-limitedpump`.
- `.build/logs/msbuild-20260510_191142_208.log` — Debug x64 `RedConfigure`
  build after fixing the compact first-open Localization overlap and increasing
  the initial window size; build wrapper diagnostics: 0 warnings, 0 errors.
- `.build/x64/Debug/RedConfigureTests.exe` on 2026-05-10 — passed after the
  compact Localization layout fix.
- `.build/logs/msbuild-20260510_182424_353.log` — Debug x64
  `RedConfigureTests` build after adding translation table view coverage for
  search, ID/status filters, and ascending/descending sort; build wrapper
  diagnostics: 0 warnings, 0 errors.
- `.build/x64/Debug/RedConfigureTests.exe` on 2026-05-10 — passed after adding
  the translation view search/filter/sort regression.
- `.build/logs/msbuild-20260510_182741_417.log` — Debug x64 `RedConfigure`
  build after replacing the two-table Localization layout with one sortable,
  searchable, filterable resource table and adding readable culture display
  names; build wrapper diagnostics: 0 warnings, 0 errors.
- `.build/logs/msbuild-20260510_175909_149.log` — Debug x64
  `RedConfigureTests` build after workspace-root normalization coverage was
  added; build wrapper diagnostics: 0 warnings, 0 errors.
- `.build/x64/Debug/RedConfigureTests.exe` on 2026-05-10 — passed after adding
  coverage for resolving `.build\x64\Debug` launch paths back to the
  repository root.
- `.build/logs/msbuild-20260510_180124_816.log` — Debug x64 `DxUiTests` build
  after the repeated `Ctrl+Backspace` path-editing regression was added; build
  wrapper diagnostics: 0 warnings, 0 errors.
- `.build/x64/Debug/DxUiTests.exe TextInputBridge` on 2026-05-10 — historical pre-retirement pass after
  fixing repeated Ctrl+Backspace path segment deletion and translated delete
  character suppression in the DxUi text bridge.
- `.build/logs/msbuild-20260510_180753_126.log` — Debug x64 `RedConfigure`
  build after moving culture/owner controls into Localization, adding official
  culture choices, root auto-normalization, richer theme samples, preview
  click-to-select, active-theme authored color keys, and first theme group batch controls; build wrapper
  diagnostics: 0 warnings, 0 errors.
- `.build/logs/msbuild-20260510_172040_185.log` — Debug x64
  `RedConfigureTests` build after task-mode navigation and theme expression
  model coverage were added; build wrapper diagnostics: 0 warnings, 0 errors.
- `.build/x64/Debug/RedConfigureTests.exe` on 2026-05-10 — passed after adding
  coverage for four RedConfigure task modes, expression-authored theme preview
  colors, cycle rejection, and flattened expression export.
- `.build/logs/msbuild-20260510_172059_204.log` — Debug x64
  `RedConfigure` build after task-mode UI, icon resource, and theme expression
  authoring changes; build wrapper diagnostics: 0 warnings, 0 errors.
- `.build/logs/msbuild-20260509_191136_532.log` — Debug x64
  `MonitorTest` build after the monitor document batch/filter guard was added;
  build wrapper diagnostics: 0 warnings, 0 errors.
- `.build/logs/msbuild-20260509_191616_355.log` — Debug x64
  `RedSalamanderMonitor` build after the deque-backed ETW queue drain,
  document batch append, filter-anchor helper usage, and stale Comctl32
  dependency removal; build wrapper diagnostics: 0 warnings, 0 errors.
- `Specs/TestRuns/4cb089111a23/Monitor/2026-05-09_191537/` —
  `RedSalamanderMonitor.exe --chrome-selftest` after test-enabled build;
  passed including the synthetic ETW batch queue drain guard
  (`before=0 after=260 expectedDelta=260`) and perf metrics for the 200 + 60
  bounded document-append batches.
- `.build/logs/msbuild-20260505_162852_589.log` — Debug x64
  `RedSalamander` build after shared result-emission helper wiring,
  runner-level emitted-once validation, and the duplicate
  `empty_directories` CompareDirectories case-name fix; build wrapper
  diagnostics: 0 warnings, 0 errors.
- `.build/logs/msbuild-20260505_180248_152.log` — Debug x64
  `RedSalamander` build after FileOperations custom speed-limit prompt,
  postmortem diagnostic, and parallel directory metadata fixes; build wrapper
  diagnostics: 0 warnings, 0 errors.
- `.build/logs/msbuild-20260505_182728_031.log` — Debug x64
  `RedSalamander` build after FileOperations phase-prefix filter support; build
  wrapper diagnostics: 0 warnings, 0 errors.
- `.build/logs/msbuild-20260505_160654_612.log` — Debug x64
  `RedSalamander` build after bounded self-test timeout parsing, stricter
  archive repo-root discovery, Preferences self-test namespace cleanup, and
  runner-native self-test case listing; build wrapper diagnostics: 0 warnings,
  0 errors.
- `.build/logs/msbuild-20260505_154051_915.log` — Debug x64 `DxUiTests`
  build after explicit unknown/empty CLI argument rejection; diagnostics: 0
  warnings, 0 errors.
- `Specs/TestRuns/4cb089111a23/Commands/2026-05-05_154805/` — focused
  `red_salamander_help_lists_diagnostics_options` Commands self-test, passed
  with 1 passed / 0 failed / 0 skipped after archive discovery began requiring
  `.git` identity.
- `Specs/TestRuns/4cb089111a23/Commands/2026-05-05_160043/` — focused
  `cmd_preferences_dialog_category_tree_uses_dxui_host_without_visible_legacy_treeview`
  Commands self-test after Preferences namespace cleanup; passed with 1 passed
  / 0 failed / 0 skipped.
- `Specs/TestRuns/4cb089111a23/Commands/2026-05-05_163030/` — focused
  `cmd_preferences_dialog_category_tree_uses_dxui_host_without_visible_legacy_treeview`
  Commands self-test through `Tools/Run-AllTests.ps1`, including runner-native
  expected/result name validation; passed with 1 passed / 0 failed / 0 skipped.
- `Specs/TestRuns/4cb089111a23/Tests/2026-05-05_163053_selftest_case_inventory/`
  — `RedSalamander.exe --selftest-list-cases` runner-native inventory emitted
  the pre-fix 817 total entries: 145 CompareDirectories, 597 Commands, and 75
  FileOperations phases, with zero duplicate case names across all suites. The
  filtered form
  `--selftest-list-cases --commands-selftest --selftest-case=cmd_preferences_`
  emitted 168 Commands cases.
- `Specs/TestRuns/4cb089111a23/FileOperations/2026-05-05_1738_prompt_async_focus_after_assertion/`
  — focused `Phase7_ParallelCopyMoveKnobs` run after async custom speed-limit
  prompt dispatch; passed with 0 failures.
- `Specs/TestRuns/4cb089111a23/FileOperations/2026-05-05_1747_phase12_phase13_focus_after_fix/Phase13_PostMortemDiagnostics/`
  — focused Phase13 run after self-contained diagnostic seeding; passed with 0
  failures.
- `Specs/TestRuns/78ac7c415c54/FileOperations/2026-05-05_1805_phase12_parallel_directory_metadata_after_fix/`
  — focused `Phase12_ReparsePointPolicy` run after deferred parallel directory
  metadata restoration; passed with 3 passed / 0 failed / 0 skipped.
- `Specs/TestRuns/78ac7c415c54/FileOperations/2026-05-05_1830_phase5_prefix_filter_pass/`
  — `Run-AllTests.ps1 -Suite FileOps -CaseFilter Phase5_` prefix-filter run;
  passed with 9 passed / 0 failed / 0 skipped.
- `Specs/TestRuns/78ac7c415c54/FileOperations/2026-05-05_1842_fileops_full_after_fixes_pass/`
  — full `Run-AllTests.ps1 -Suite FileOps -SkipBuild -TimeoutMultiplier 2`
  recheck after the FileOperations fixes; passed with 55 passed / 0 failed / 20
  expected skips.
- `Specs/TestRuns/78ac7c415c54/FileOperations/2026-05-05_1914_phase7_speed_prompt_blocker_regression_pass/`
  — focused `Phase7_ParallelCopyMoveKnobs` recheck for the custom speed-limit
  prompt blocker; passed with 3 passed / 0 failed / 0 skipped and no lingering
  prompt process.
- `Specs/TestRuns/4cb089111a23/FileOps/2026-05-05_203552/` — focused
  `Phase14_PopupHostLifetimeGuard` recheck through `Run-AllTests.ps1` with
  `REDSALAMANDER_SELFTEST_ROOT` pointed at an isolated worktree-local root;
  passed with 3 passed / 0 failed / 0 skipped and proved the runner reads the
  same isolated `last_run` tree as the native harness.
- `.build/logs/msbuild-20260505_212109_674.log` — Debug x64 full build after
  the Connection Manager scrollable-editor pointer-toggle fix; build wrapper
  diagnostics: 0 warnings, 0 errors.
- `Specs/TestRuns/4cb089111a23/Commands/2026-05-05_212251/` — focused
  `cmd_connection_manager_window_pointer_click_toggles_visible_dx_toggle`
  recheck after scrolling the S3 `Use HTTPS` toggle into view and mapping the
  debug rectangle through the editor `ScrollPanel` offset; passed with 1 passed
  / 0 failed / 0 skipped. `selftest_run_trace.txt` records the toggle rect
  moving from the failing offscreen content-space range to client
  `(564,576)-(1050,612)` before the two real pointer clicks.
- `Specs/TestRuns/4cb089111a23/Commands/2026-05-05_212506/`,
  `2026-05-05_212508/`, `2026-05-05_212512/`, and
  `2026-05-05_212513/` — adjacent pointer-toggle rechecks after the
  Connection Manager scrollable-editor fix; Plugin Configuration `Use HTTPS`,
  credential prompt secret visibility, Compare Options, and Find recursive
  checkbox each passed with 1 passed / 0 failed / 0 skipped.
- `Specs/TestRuns/4cb089111a23/Commands/2026-05-05_212659/` — full
  `cmd_connection_manager_window_` family recheck after the scrollable-editor
  pointer fix; passed with 29 passed / 0 failed / 0 skipped.
- `.build/logs/msbuild-20260505_200710_000.log` — Debug x64 full build after
  the Compare/Search service no-wait, direct-SQLite precondition, isolated
  foreground-store, and request-budget fixes; build wrapper diagnostics: 0
  warnings, 0 errors.
- `Specs/TestRuns/78ac7c415c54/CompareDirectories/2026-05-05_2012_compare_full_after_service_cleanup_pass/`
  — full `Run-AllTests.ps1 -Suite Compare -SkipBuild -TimeoutMultiplier 2`
  recheck after the Compare/Search cleanup; passed with 125 passed / 0 failed /
  24 expected skips. The direct-SQLite-only service failure/prefilter cases
  skipped on this machine because the live NTFS journal cursor was unavailable.
- `.build/logs/msbuild-20260505_224017_123.log` — Debug x64 build after
  Commands Compare Options pane-settle guards; build wrapper diagnostics: 0
  warnings, 0 errors.
- `.build/logs/msbuild-20260505_224402_884.log` — Debug x64 build after
  app-side `REDSALAMANDER_SELFTEST_ROOT` absolute normalization; build wrapper
  diagnostics: 0 warnings, 0 errors.
- `Specs/TestRuns/4cb089111a23/Commands/2026-05-05_224215/` — focused
  `cmd_compare_directories_options_live_dx_body_interaction` regression run
  before app-side self-test root normalization; failed quickly with a bounded
  pane-settle diagnostic instead of hanging on stale `C:\` pane roots.
- `Specs/TestRuns/4cb089111a23/Commands/2026-05-05_225225/` — same focused
  Compare Options live-DxUi case after absolute self-test root normalization;
  passed with 1 passed / 0 failed / 0 skipped, but exposed an unacceptably slow
  405 s runtime and a 227 MB perf log from unbounded Commands message pumping.
- `.build/logs/msbuild-20260505_225908_817.log` — Debug x64 full build after
  bounding the shared Commands message pump; build wrapper diagnostics: 0
  warnings, 0 errors.
- `Specs/TestRuns/4cb089111a23/Commands/2026-05-05_230056/` — focused
  `cmd_compare_directories_options_live_dx_body_interaction` recheck after the
  bounded message pump; passed with 1 passed / 0 failed / 0 skipped in about 5
  seconds and reduced the perf log to about 1.7 MB.
- `.build/logs/msbuild-20260506_084509_961.log` — Debug x64 build after the
  Compare Options retained-focus restoration and post-wait diagnostic fix;
  build wrapper diagnostics: 0 warnings, 0 errors.
- `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_084721/`,
  `2026-05-06_084755/`, and `2026-05-06_084811/` — three fresh-root foreground
  Compare Options family repeats after the retained-focus restoration; each
  passed with 9 passed / 0 failed / 0 skipped. The traversal trace records the
  formerly failing reverse step as `afterFocus=7 afterScroll=162/312`.
- `Invoke-Pester .\Tools\Tests -ExcludeTag RequiresBuildToolchain -PassThru`
  on 2026-05-05 — 47 artifact-safe tooling tests passed, 0 failed; the single
  excluded test remains tagged `RequiresBuildToolchain`.

## Current Coverage Challenge Notes (2026-05-05)

The 2026-05-05 test review challenged the named gap areas against the current
code before adding duplicate tests. Treat this section as the durable summary of
what is covered now, what remains intentionally blocked, and which cases are
correctness-only rather than performance gates.

Commands coverage resolved by audit:

- `Commands.SelfTest.ShellCommands.cpp` covers focused-item Security shell
  action routing, current-directory context-menu routing, NTFS alternate data
  stream removal, recursive Change Attributes progress, shortcut/link targets,
  junction navigation, Shell New templates, and clipboard shortcut paths.
- Plugin Configuration dialog tests cover visible UIA provider counts,
  Value/Toggle/Invoke pattern exposure, accessible names, live DxUi edit
  mutation, Browse cancel behavior, OK/Cancel invocation through UIA, tab
  traversal, access keys, pointer toggles, theme cycling, and long-run
  open/close/scroll stability.
- Connection Manager pointer-toggle coverage now exercises the S3 `Use HTTPS`
  toggle through the real pointer path after explicitly scrolling the
  scrollable editor field into the viewport; debug rectangles used by that
  test are client-space rectangles after `ScrollPanel` offset is applied.
- Compare Options live-DxUi interaction coverage now prepares absolute,
  settled left/right pane roots before opening Compare Directories and relies
  on a bounded Commands message pump so hover/cursor traffic cannot mask the
  duration of the actual UI interaction steps. Its tab-traversal case covers
  focus changes that scroll the retained DxUi body and verifies that retained
  body focus survives the resulting layout pass.
- Compare Options and Plugin Configuration cancel-persistence paths already
  verify that live DxUi edits/toggles are discarded on cancel and preserved only
  on accepted save paths.
- Connection Manager covers clean external reload refresh and dirty/stale save
  prompts while open. It does not simulate a remote provider refreshing an
  already-open secret; that belongs with deterministic provider failure seams.
- No implemented Commands long-poll product contract was found. Do not add a
  synthetic long-poll case until a durable product requirement names the loop,
  cadence, and user-visible failure mode.

CompareDirectories remaining gaps:

- Covered today: OAuth auth-mode and refresh-token storage/deletion/session
  behavior, Google Drive missing-refresh-token gates, SQLite bootstrap,
  compaction, WAL checkpoint/truncation, legacy upgrade, future-schema
  rejection, content compare equal/different/no-I/O/Unicode/short-read paths,
  synthetic crash-quarantine markers, bounded cancellation, service no-wait
  live-scan degradation when SQLite is not ready/current, service-vs-host
  warning semantics, startup fixed-root discovery with isolated stores, and
  rebuild control through a deterministic snapshot-backed foreground service.
- Blocked until deterministic seams exist: expired/revoked OAuth refresh,
  network failure during refresh, HTTP 429/retry-after backoff, hard
  content-compare read failures such as sharing violation or access denied, and
  true search-service process death mid-query.
- Partial future breadth: a narrowly timed cancel during a content-compare
  stamp update is not isolated from the broader bounded-cancel and pending
  content-comparison coverage.

FileOperations remaining gaps:

- Covered today: pre-calc cancel, queued cancel, local bandwidth cancel
  latency, active-worker cancellation in the recursive matrix, bridge
  throughput/concurrency, conflict prompts, async custom speed-limit prompts,
  settings defaults, connection overrides, directory metadata preservation in
  recursive copy, postmortem diagnostic export, and junction/reparse policy
  including protected junctions, out-of-tree junctions, copied reparse tags,
  move rollback, bridge copy/move of reparse items, and reparse delete target
  preservation.
- Blocked until deterministic OS/test-environment seams exist: disk-full or
  quota failures, ACL-denied copy destinations with issues-pane assertions, UNC
  destination fixtures, and cloud-placeholder recall behavior.
- Partial future breadth: mid-bridge cancel/rollback, concurrent destination
  rename races, relative symlink behavior, and runtime bandwidth/parallelism
  settings mutation after data transfer begins.

Standalone and DxUi coverage notes:

- DxUiTests already covers typeahead, scrollbars, single-line editing,
  native TextField routing, ComboBox popup scrolling/hover, Tree
  selection/context/toggle, reduced motion, and UIA provider patterns. Future
  additions should target controls with only visual-baseline coverage.
- MonitorTest is intentionally narrow: ETW TraceLogging emit/receive,
  self-diagnostic suppression, invalid-rectangle opt-in, self-originated
  event rejection, and the ColorTextView scrollbar visibility model are
  covered, but the full Monitor window state machine and a rich ETW
  filtering-rule matrix are not.
- LocalizationTests covers embedded fallback, satellite string/menu/dialog
  lookup, localized dialog templates, invalid-culture fallback, and persisted
  `ui.language` roundtrips. It does not enumerate every shipped satellite.
- PerformanceTests2 covers 14 CppUnitTest cases focused on icon enumeration,
  duplicate-path refresh/compact-mode hit testing, FolderView column layout and
  sort threshold policy, splash close guard, and empty plugin-manager discovery;
  it is not a general rendering/search/Compare throughput suite.

Self-test metric-family map:

- Commands self-test metrics currently come from ViewerText diff/hex baselines:
  `viewer.diff.*` open-to-first-visible, semantic paint, visible row counts,
  theme switch, scroll repaint, hunk jump, context expansion, viewport
  rehydration/backtrack, referenced-byte growth, deferred rows, and placeholder
  rows/bands.
- CompareDirectories self-test metrics currently cover local wide-tree search:
  `compare.selftest.local_search_scan_wide_tree_workers_us` and
  `compare.selftest.local_search_scan_wide_tree_us`. The Commands case
  `cmd_compare_directories_progress_perf` is a progress correctness/stability
  guard today; it must emit a self-test-local metric before being used as a
  performance gate.
- FileOperations self-test metrics currently cover pre-calc cancel latency,
  local and parallel bandwidth throttling, copy/move concurrency, auto
  concurrency, recursive copy matrices, recycle-bin batching, default bandwidth
  limits, conflict convergence, bridge pipeline, connection overrides, and the
  global connection gate.
- No new performance instrumentation was added by the 2026-05-05 test review
  helper/listing/reporting changes, so no before/after `perf_metrics.jsonl`
  archive is required for those changes.

Recent UI-retirement evidence:

- `Z:\src\RedSalamander\.build\logs\msbuild-20260517_195623_555.log` — `DxUiTests` Debug build after deleting the remaining production `TextInputBridge` surface from `Common/DxUi`, removing the retired bridge window messages and audit allowlist rows, and trimming bridge-suite-only test helpers; diagnostics `0 warning(s), 0 error(s)`. Serial `NativeTextInput`, `MultilineText`, `ReadOnly`, `TextField`, `ComboBox`, `Accessibility`, and `WindowHost` suites exited 0 with perf rows in `Specs/TestRuns/local_scratch/dxui_native_textinput_after_production_bridge_removal_20260517_1958.jsonl`, `dxui_multiline_after_production_bridge_removal_20260517_1958.jsonl`, `dxui_readonly_after_production_bridge_removal_20260517_1958.jsonl`, `dxui_textfield_after_production_bridge_removal_20260517_1958.jsonl`, `dxui_combobox_after_production_bridge_removal_20260517_1958.jsonl`, `dxui_accessibility_after_production_bridge_removal_20260517_1958.jsonl`, and `dxui_windowhost_after_production_bridge_removal_20260517_1958.jsonl`, leaving 628 registered DxUi component tests and no production bridge code. `Audit-RemainingWin32UiDependencies.ps1 -FailOnFindings`, `Audit-VisibleNativeSurfaces.ps1`, `Audit-ComctlReportSurfaces.ps1`, and the production bridge-name audit all exited 0; `git diff --check` exited 0 with only LF/CRLF conversion warnings.
- `Specs/TestRuns/4cb089111a23/Audit/2026-05-17_2005_dxui_native_textinput_bridge_removed_final_recheck/` — refreshed broad/visible/comctl audit bundle after production bridge deletion and authoritative spec cleanup; `Audit-RemainingWin32UiDependencies.ps1 -FailOnFindings`, `Audit-VisibleNativeSurfaces.ps1`, and `Audit-ComctlReportSurfaces.ps1` each exited 0 with zero unallowlisted findings.
- `Z:\src\RedSalamander\.build\logs\msbuild-20260517_194303_147.log` — `DxUiTests` Debug build after deleting the final `TextInputBridge` compatibility suite source, runner registration, and project entry because retained/native semantic coverage now lives in `NativeTextInput`, `MultilineText`, and `ReadOnly`; diagnostics `0 warning(s), 0 error(s)`. Serial replacement suites exited 0 with perf rows in `Specs/TestRuns/local_scratch/dxui_native_textinput_after_textinputbridge_suite_removal_20260517_1943.jsonl`, `Specs/TestRuns/local_scratch/dxui_multiline_after_textinputbridge_suite_removal_20260517_1944.jsonl`, and `Specs/TestRuns/local_scratch/dxui_readonly_after_textinputbridge_suite_removal_20260517_1944.jsonl`, leaving 629 registered DxUi component tests and no runnable `TextInputBridge` suite.
- `Z:\src\RedSalamander\.build\logs\msbuild-20260517_193902_085.log` — `DxUiTests` Debug build after adding native host-HWND focus-loss session teardown/regain plus multiline/wrapped Return default-button suppression, Tab/Shift+Tab traversal, and Escape cancel routing, then retiring twelve hidden-child `WM_KILLFOCUS` / bridge special-key dialog-flow probes from `DxUiTests.TextInputBridge.cpp`; diagnostics `0 warning(s), 0 error(s)`. Serial `NativeTextInput` and `TextInputBridge` exited 0, leaving 117 bridge compatibility cases and raising `NativeTextInput` to 109 cases, with perf rows in `Specs/TestRuns/local_scratch/dxui_native_textinput_multiline_dialog_bridge_retirement_20260517_1940.jsonl` and `Specs/TestRuns/local_scratch/dxui_textinputbridge_after_multiline_dialog_retirement_20260517_1940.jsonl`.
- `Z:\src\RedSalamander\.build\logs\msbuild-20260517_183645_412.log` — `DxUiTests` Debug build after adding native multiline/wrapped host-HWND character replacement plus Return insertion/replacement state coverage, then retiring twelve hidden-edit `WM_SETTEXT` / `EM_REPLACESEL` / `WM_CHAR` / Return bridge probes from `DxUiTests.TextInputBridge.cpp`; diagnostics `0 warning(s), 0 error(s)`. Serial `NativeTextInput` and `TextInputBridge` exited 0, leaving 129 bridge compatibility cases and raising `NativeTextInput` to 107 cases, with perf rows in `Specs/TestRuns/local_scratch/dxui_native_textinput_multiline_edit_probe_retirement_20260517_1837.jsonl` and `Specs/TestRuns/local_scratch/dxui_textinputbridge_after_multiline_edit_probe_retirement_20260517_1837.jsonl`.
- `Z:\src\RedSalamander\.build\logs\msbuild-20260517_183037_025.log` — `DxUiTests` Debug build after adding native multiline/wrapped IME composition/candidate anchoring for caret moves across logical/visual lines and focused-control moves, then retiring six hidden-edit IMM32 placement variants from `DxUiTests.TextInputBridge.cpp`; diagnostics `0 warning(s), 0 error(s)`. Serial `NativeTextInput` and `TextInputBridge` exited 0, leaving 141 bridge compatibility cases and raising `NativeTextInput` to 106 cases, with perf rows in `Specs/TestRuns/local_scratch/dxui_native_textinput_ime_window_bridge_retirement_20260517_1831.jsonl` and `Specs/TestRuns/local_scratch/dxui_textinputbridge_after_ime_window_retirement_20260517_1832.jsonl`.
- `Z:\src\RedSalamander\.build\logs\msbuild-20260517_182322_471.log` — `DxUiTests` Debug build after adding native single-line/multiline/wrapped IME result-only host-key resume coverage plus result-and-continuing-composition key ownership coverage, then retiring six hidden-bridge IME result-routing variants from `DxUiTests.TextInputBridge.cpp`; diagnostics `0 warning(s), 0 error(s)`. The first serial `NativeTextInput` attempt with `dxui_native_textinput_ime_result_routing_bridge_retirement_20260517_1822.jsonl` exposed an over-strict test assertion for multiline `WM_KEYDOWN/VK_RETURN` text insertion, and the corrected routing-only assertion then passed. Final serial `NativeTextInput` and `TextInputBridge` exited 0, leaving 147 bridge compatibility cases and raising `NativeTextInput` to 104 cases, with perf rows in `Specs/TestRuns/local_scratch/dxui_native_textinput_ime_result_routing_bridge_retirement_20260517_1823.jsonl` and `Specs/TestRuns/local_scratch/dxui_textinputbridge_after_ime_result_routing_retirement_20260517_1824.jsonl`.
- `Z:\src\RedSalamander\.build\logs\msbuild-20260517_181748_496.log` — `DxUiTests` Debug build after adding native multiline/wrapped multiline IME composition-owned Return/Escape/Tab coverage and retiring three hidden-bridge IME special-key variants from `DxUiTests.TextInputBridge.cpp`; diagnostics `0 warning(s), 0 error(s)`. Serial `NativeTextInput` and `TextInputBridge` exited 0, leaving 153 bridge compatibility cases and raising `NativeTextInput` to 102 cases, with perf rows in `Specs/TestRuns/local_scratch/dxui_native_textinput_ime_special_key_bridge_retirement_20260517_1818.jsonl` and `Specs/TestRuns/local_scratch/dxui_textinputbridge_after_ime_special_key_retirement_20260517_1818.jsonl`.
- `Z:\src\RedSalamander\.build\logs\msbuild-20260517_181403_228.log` — `DxUiTests` Debug build after adding native single-line tab-character suppression plus partial-selection paste state coverage and retiring eight hidden-bridge single-line edit-routing probes from `DxUiTests.TextInputBridge.cpp`; diagnostics `0 warning(s), 0 error(s)`. Serial `NativeTextInput` and `TextInputBridge` exited 0, leaving 156 bridge compatibility cases and raising `NativeTextInput` to 101 cases, with perf rows in `Specs/TestRuns/local_scratch/dxui_native_textinput_singleline_bridge_probe_retirement_20260517_1814.jsonl` and `Specs/TestRuns/local_scratch/dxui_textinputbridge_after_singleline_probe_retirement_20260517_1815.jsonl`.
- `Z:\src\RedSalamander\.build\logs\msbuild-20260517_180757_053.log` — `DxUiTests` Debug build after adding native editable ComboBox exact-match selection, undo/redo, paste, Delete, Ctrl+Delete, Ctrl+A replacement, and path-style Ctrl+Backspace coverage, then retiring four hidden-bridge editable ComboBox duplicate tests from `DxUiTests.TextInputBridge.cpp`; diagnostics `0 warning(s), 0 error(s)`. Serial `NativeTextInput` and `TextInputBridge` exited 0, leaving 164 bridge compatibility cases and raising `NativeTextInput` to 100 cases, with perf rows in `Specs/TestRuns/local_scratch/dxui_native_textinput_editable_combo_bridge_retirement_20260517_1808.jsonl` and `Specs/TestRuns/local_scratch/dxui_textinputbridge_after_editable_combo_test_retirement_20260517_1808.jsonl`.
- `Z:\src\RedSalamander\.build\logs\msbuild-20260517_180329_063.log` — `DxUiTests` Debug build after strengthening native masked retained-text character-input assertions and retiring the two hidden-bridge masked copy/cut and character-input duplicate tests from `DxUiTests.TextInputBridge.cpp`; diagnostics `0 warning(s), 0 error(s)`. Serial `NativeTextInput` and `TextInputBridge` exited 0, leaving 168 bridge compatibility cases, with perf rows in `Specs/TestRuns/local_scratch/dxui_native_textinput_after_masked_bridge_retirement_20260517_1803.jsonl` and `Specs/TestRuns/local_scratch/dxui_textinputbridge_after_masked_test_retirement_20260517_1803.jsonl`.
- `Z:\src\RedSalamander\.build\logs\msbuild-20260517_180026_689.log` — `DxUiTests` Debug build after retiring the two single-line hidden-edit word-selection comparison tests from `DxUiTests.TextInputBridge.cpp`; diagnostics `0 warning(s), 0 error(s)`. Serial `NativeTextInput` and `TextInputBridge` exited 0, leaving 170 bridge compatibility cases, with perf rows in `Specs/TestRuns/local_scratch/dxui_native_textinput_after_word_selection_bridge_retirement_20260517_1800.jsonl` and `Specs/TestRuns/local_scratch/dxui_textinputbridge_after_word_selection_test_retirement_20260517_1801.jsonl`.
- `Z:\src\RedSalamander\.build\logs\msbuild-20260517_175718_278.log` — `DxUiTests` Debug build after moving the multiline/wrapped context-menu anchor assertion onto native host-HWND `VK_APPS` / `Shift+F10` coverage and retiring the four hidden-bridge duplicate context-menu tests from `DxUiTests.TextInputBridge.cpp`; diagnostics `0 warning(s), 0 error(s)`. Serial `NativeTextInput` and `TextInputBridge` exited 0, leaving 172 bridge compatibility cases, with perf rows in `Specs/TestRuns/local_scratch/dxui_native_textinput_context_menu_bridge_retirement_20260517_1757.jsonl` and `Specs/TestRuns/local_scratch/dxui_textinputbridge_after_context_menu_test_retirement_20260517_1758.jsonl`.
- `Z:\src\RedSalamander\.build\logs\msbuild-20260517_175236_429.log` — `DxUiTests` Debug build after retiring four bridge-only hidden-window detail tests from `DxUiTests.TextInputBridge.cpp`; diagnostics `0 warning(s), 0 error(s)`. Serial `TextInputBridge` exited 0 with 176 remaining compatibility cases and perf rows in `Specs/TestRuns/local_scratch/dxui_textinputbridge_after_hidden_detail_test_retirement_20260517_1752.jsonl`.
- `Z:\src\RedSalamander\.build\logs\msbuild-20260517_174859_672.log` — `DxUiTests` Debug build after the native text-input command-line final recheck; diagnostics `0 warning(s), 0 error(s)`. Serial `NativeTextInput`, `TextField`, and `ComboBox` rechecks exited 0 with perf rows in `Specs/TestRuns/local_scratch/dxui_native_textinput_current_recheck_20260517_1749.jsonl`, `dxui_textfield_current_recheck_20260517_1750.jsonl`, and `dxui_combobox_current_recheck_20260517_1750.jsonl`.
- `Specs/TestRuns/4cb089111a23/Commands/2026-05-17_174604/` — final focused `cmd_pane_command_line_insertion_and_execute` command selftest after replacing the FolderWindow command-line visible native `STATIC` / `EDIT` controls with a DxUi native-backend `TextField`; launched with `Start-Process -Wait -PassThru`; 1 passed, 0 failed. Green candidate evidence for the same guard is `Specs/TestRuns/4cb089111a23/Commands/2026-05-17_173548/`; red evidence is `Specs/TestRuns/4cb089111a23/Commands/2026-05-17_172957/`, which failed at the new DxUi-host assertion before the migration.
- `Specs/TestRuns/4cb089111a23/Audit/2026-05-17_1748_dxui_native_textinput_commandline_final_recheck/` — final broad remaining Win32 UI dependency audit after command-line migration and the test-only FolderView synthetic-thumbnail allowlist refresh; `HDC text/selection bridge` 16 total / 16 allowed / 0 unallowed, `HFONT handle` 1 total / 1 allowed / 0 unallowed, `Native visible control creation` 1 total / 1 allowed / 0 unallowed. The same archive includes clean `Audit-VisibleNativeSurfaces.ps1` and `Audit-ComctlReportSurfaces.ps1` outputs.
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-29_111643/` — Connection Manager battle-test family after dialog shim retirement; `cmd_connection_manager_window_` passed 25 passed / 0 failed / 0 skipped and archived `perf/perf_metrics.jsonl`.
- `Specs/TestRuns/4cb089111a23/Audit/2026-04-26_012857_remaining_win32_ui_dependency_post_closeout_recheck/` — post-closeout broad HFONT/GDI/native-control audit recheck. Audit perf: 3.446 seconds; zero unallowlisted findings.
- `Specs/TestRuns/4cb089111a23/DxUiTests/2026-04-26_012857_remaining_win32_ui_dependency_post_closeout_menu/` — post-closeout focused DxUi menu run, exited 0 in 15.373 seconds. Current focused invocation syntax is `.build\x64\Debug\DxUiTests.exe --suite=Menu`.
- `Specs/TestRuns/4cb089111a23/Audit/2026-04-26_012017_remaining_win32_ui_dependency_final_recheck/` — final wrapper broad HFONT/GDI/native-control audit recheck. Audit perf: 3.315 seconds; zero unallowlisted findings.
- `.build/logs/msbuild-20260426_010339_544.log` and `.build/logs/msbuild-20260426_010518_794.log` — wrapper-pass Debug/Release builds after the Function Bar paint-DC stabilization and final closeout documentation refresh.
- `Specs/TestRuns/4cb089111a23/DxUiTests/2026-04-26_011821_remaining_win32_ui_dependency_final_full_dxui/` — full `.build\x64\Debug\DxUiTests.exe` run, exited 0 in 110.547 seconds.
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-26_010657/`, `2026-04-26_010717/`, `2026-04-26_010746/`, `2026-04-26_011225_connection_manager_final_recheck/`, `2026-04-26_011255_plugin_configuration_final_recheck/`, and `2026-04-26_011808_preferences_final_recheck/` — wrapper focused command recheck for UI chrome, status bar shell stability, Compare Options, Connection Manager, Plugin Configuration, and Preferences; all exited 0 with zero failures.
- `Specs/TestRuns/4cb089111a23/Audit/2026-04-26_005800_remaining_win32_ui_dependency_final_recheck/` — final broad HFONT/GDI/native-control audit recheck. Audit perf: 3.352 seconds; zero unallowlisted findings.
- `.build/logs/msbuild-20260426_005323_200.log` and `.build/logs/msbuild-20260426_005751_693.log` — final Debug/Release `RedSalamander` builds after the Function Bar layout/material-tolerant pixel guard.
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-26_005455/`, `2026-04-26_005540/`, and `2026-04-26_005709/` — final focused command recheck for UI chrome, status bar shell stability, and Compare Options; all passed with zero failures. The failed `2026-04-26_005614/` Compare Options attempt is non-gating because the same failed case passed in isolation at `2026-04-26_005634/` and the full family passed on rerun.
- `Specs/TestRuns/4cb089111a23/DxUiTests/2026-04-26_005914_remaining_win32_ui_dependency_final_menu/` — final focused DxUi menu recheck, exited 0 in 14.632 seconds. Current focused invocation syntax is `DxUiTests.exe --suite=Menu`.
- `Specs/TestRuns/4cb089111a23/Audit/2026-04-26_001034_remaining_win32_ui_dependency_recheck/` — current-day broad HFONT/GDI/native-control audit recheck. Audit perf: 3.339 seconds; zero unallowlisted findings.
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-26_001212/`, `2026-04-26_001216/`, and `2026-04-26_001234/` — current-day focused command recheck for UI chrome, status bar shell stability, and Compare Options; all passed with zero failures. The failed `2026-04-26_001047/` and `2026-04-26_001156/` `cmd_app_toggleUiChrome` attempts are non-gating because they launched the visibility-sensitive GUI selftest with a hidden top-level window.
- `Specs/TestRuns/4cb089111a23/DxUiTests/2026-04-26_001241_remaining_win32_ui_dependency_recheck_menu/` — current-day focused DxUi menu recheck, exited 0 in 15.259 seconds. Current focused invocation syntax is `DxUiTests.exe --suite=Menu`.
- `Specs/TestRuns/4cb089111a23/Audit/2026-04-25_235110_remaining_win32_ui_dependency_closeout/` — broad HFONT/GDI/native-control audit after the remaining app-owned HDC paint/selection seams moved behind `D2DHdcPaint::Session` and the last Compare Options legacy `STATIC` fallback was replaced. Audit perf: 3.219 seconds; zero unallowlisted findings.
- `.build/logs/msbuild-20260425_234924_528.log` — Debug build after the remaining paint/static closeout.
- `.build/logs/msbuild-20260425_235255_302.log` — Release build after the remaining paint/static closeout.
- `.build/logs/msbuild-20260425_235401_128.log` and `Specs/TestRuns/4cb089111a23/DxUiTests/2026-04-26_000537_remaining_win32_ui_dependency_closeout/` — DxUiTests build plus full DxUiTests run, exited 0.
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_235651/`, `2026-04-25_235730/`, `2026-04-25_235809/`, `2026-04-25_235945/`, `2026-04-26_000025/`, and `2026-04-26_000514/` — closeout focused command refresh for UI chrome, status bar, Compare Options, Connection Manager, Plugin Configuration, and Preferences; all passed with zero failures.
- `Specs/TestRuns/4cb089111a23/Audit/2026-04-25_232753_remaining_win32_ui_native_hosts_reduced_candidate/` — broad HFONT/GDI/native-control audit after shared native-font helper removal and DxUi host-window class cleanup. Audit perf: 3.433 seconds; remaining closeout blockers are 22 HDC text/selection bridges and one Compare Options legacy `STATIC` fallback.
- `.build/logs/msbuild-20260425_232017_534.log` — Debug build after plugin viewer unused native font handles were removed.
- `.build/logs/msbuild-20260425_232609_620.log` — Debug build after native host cleanup.
- `.build/logs/msbuild-20260425_233112_306.log` — Release build after native host cleanup.
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_232948/` — Connection Manager focused family, 13 passed, 0 failed.
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_233010/` — Plugin Configuration focused family, 10 passed, 0 failed.
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_233026/` — Compare Options focused family, 9 passed, 0 failed.
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_233028/`, `2026-04-25_233032/`, and `2026-04-25_233037/` — Preferences shell/general/panes refresh after shell font propagation removal.
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_233054/` — Navigation edit-suggest keyboard routing direct focused rerun passed after hidden-launch archive `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_233044/` failed transiently.
- `Specs/TestRuns/4cb089111a23/Audit/2026-04-25_224400_remaining_win32_ui_bridge_allowlist_uimetrics_viewerweb_candidate/` — broad HFONT/GDI/native-control audit after hidden text bridge allowlist, `Win32UiHelpers` deletion/`UiMetrics` split, dead HFONT measurement bridge removal, dialog base-`LOGFONT` cloning removal, and ViewerWeb DirectWrite status drawing. Audit perf: 3.331 seconds.
- `.build/logs/msbuild-20260425_222825_786.log` — Debug build after the helper split/native-font bridge cleanup.
- `.build/logs/msbuild-20260425_224509_797.log` — Release build after the helper split/native-font bridge cleanup.
- `.build\x64\Debug\DxUiTests.exe --suite=TextInputBridge` — historical focused hidden text-service bridge regression before the suite was retired, exited 0.
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_224155/` — Connection Manager focused family, 13 passed, 0 failed.
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_224242/` — Plugin Configuration focused family, 10 passed, 0 failed.
- `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_224312/`, `2026-04-25_224320/`, `2026-04-25_224331/`, and `2026-04-25_224338/` — Preferences shell/page-host/general/panes focused guards.
- Non-closeout caveat: `Specs/TestRuns/4cb089111a23/Commands/2026-04-25_224026/` was an accidental broad unfiltered Commands run and failed 4 Preferences Viewers reorder/search cases; do not use it as green closeout evidence for `Specs/Plans/Done/UI_RemainingWin32UiDependencyRetirementPlan.md`.

---

## 1. Commands Suite (`--commands-selftest`)

**Source:** `RedSalamander\SelfTest\Commands\Commands.SelfTest.cpp` orchestrator + 13 included `.cpp` family files (809 runner-listed cases; 707 static `SelfTest::RunCase` call sites)

The Commands suite is split into logical `.cpp` family files included from the main orchestrator:
- `SelfTest\Commands\Commands.SelfTest.Settings.cpp` — Settings hot-reload, store, registry, preview/file-action guards, shortcut defaults (13+ cases)
- `SelfTest\Commands\Commands.SelfTest.BatchRename.cpp` — Batch Rename dialog, rule editing, execution, report, cancel, rollback, date/time, path-sort, partial-failure coverage, gated-provider non-blocking collection, one-listing-per-parent preview validation metrics, and 10,000-row non-ASCII hashed duplicate detection
- `SelfTest\Commands\Commands.SelfTest.PluginConfig.cpp` — Plugin configuration and file-system plugin (13 cases)
- `SelfTest\Commands\Commands.SelfTest.Connections.cpp` — Connection manager and credentials (44 cases), including session-first rotation/tombstone behavior when credential persistence fails
- `SelfTest\Commands\Commands.SelfTest.Preferences.cpp` coordinator + 7 included chunk files — Preferences dialog automation (129 cases)
- `SelfTest\Commands\Commands.SelfTest.CompareOptions.cpp` — Compare directories options, chrome, and progress (15 cases)
- `SelfTest\Commands\Commands.SelfTest.Search.cpp` — Find dialog, local search index, quick search/filter (52 cases)
- `SelfTest\Commands\Commands.SelfTest.Shortcuts.cpp` — Shortcuts window (31 cases)
- `SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp` — View commands, selection, sort, pane, tabs, FolderView rendering-alert persistence, draw-item brush-reuse guards, refresh-to-paint telemetry, IconCache live-failure retry, and DPI repaint coverage (109 static registrations)
- `SelfTest\Commands\Commands.SelfTest.FileOps.cpp` — File operations issues pane, popup, and speed limit prompt (24 cases)
- `SelfTest\Commands\Commands.SelfTest.Navigation.cpp` — Navigation location, GoTo, navigation/drive menu shell stability, Command Shell Windows Terminal launch planning, Escape focus reclaim to the active FolderView, navigation-menu `Go to >` placement before drive rows, nonstandard file-system `Common Folders` submenu coverage, and directory-impact selection preservation
- `SelfTest\Commands\Commands.SelfTest.ShellCommands.cpp` — Shell-integrated pane commands including Change Attributes attributes/date-time/stream reports and recursive progress
- `SelfTest\Commands\Commands.SelfTest.Dialogs.cpp` — About, fatal error, splash, change case, filter, rename, including long initial rename-selection clipping, etc. (61 cases)

The suitetests UI automation, dialog interactions, preferences, shortcuts,
themes, navigation, and command dispatch. All cases run inside the live application
window using UIAutomation and direct Win32 message simulation.

### 1.1 Application-Level Commands (20 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `cmd_app_about_access_key_routes_ok` | About dialog keyboard access keys |
| `cmd_app_about_enter_and_escape_route_ok` | About dialog Enter/Escape handling |
| `cmd_app_about_live_dx_interaction` | About dialog DxUi surface interaction |
| `cmd_app_about_long_run_open_close_stays_stable` | About dialog stability under repeated open/close |
| `cmd_app_about_uses_dxui_surface` | About dialog uses DxUi rendering |
| `cmd_app_fatal_error_access_key_routes_ok` | Fatal error dialog access keys |
| `cmd_app_fatal_error_enter_and_escape_route_ok` | Fatal error Enter/Escape |
| `cmd_app_fatal_error_live_dx_interaction` | Fatal error DxUi interaction |
| `cmd_app_fatal_error_long_run_open_close_stays_stable` | Fatal error stability |
| `cmd_app_fatal_error_long_run_scrolling_stays_bounded` | Fatal error scroll bounds |
| `cmd_app_fatal_error_theme_matrix_keeps_message_legible` | Fatal error theme legibility |
| `cmd_app_fatal_error_uses_dxui_surface` | Fatal error DxUi surface |
| `cmd_app_fullScreen` | Full-screen toggle |
| `host_pane_alert_cookie_routes_non_focused_pane` | Host pane alert cookies route non-focused pane alerts and clears |
| `cmd_app_menuBar_mouse_open_keeps_popup_selection_clear` | Mouse-opened menu popups keep item selection empty until pointer movement or keyboard navigation |
| `cmd_app_menuBar_mouse_opened_popup_processes_keyboard_before_mouse_move` | Mouse-opened menu popups process and repaint keyboard navigation before any later mouse movement |
| `cmd_app_menuBar_persistent_direct_hover_switches_top_level_popup` | Persistent menu-bar mouse-opened popup switches directly to a neighboring top-level menu without requiring popup-item hover first |
| `cmd_app_menuBar_top_level_highlight_follows_keyboard_opened_root` | Menu bar highlight follows the active keyboard-opened root popup and does not fall back to a stale top-level hover while the popup owns focus |
| `cmd_app_menuBar_persistent_view_to_files_hover_highlight_follows_pointer` | Persistent main menu moves the top-level highlight immediately when hovering from View to Files while the View popup is open |
| `cmd_app_menuBar_persistent_view_to_plugins_hover_switches_popup` | Persistent main menu switches directly from View to Plugins on top-level hover without requiring popup-item hover first |
| `cmd_app_menuBar_submenu_placement_matches_spec` | Menu submenu placement and parent-hover retention |
| `cmd_app_swapPanes` | Pane swap command |
| `cmd_app_toggleUiChrome` | UI chrome visibility toggle, function bar layout, painted themed pixels, DirectWrite text metrics, and splitter arrow affordances |

### 1.2 Prompt and Splash (5 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `cmd_app_prompt_access_keys_route_expected_actions` | Prompt dialog access keys |
| `cmd_app_prompt_long_run_open_close_stays_stable` | Prompt dialog stability |
| `cmd_app_prompt_uses_alert_overlay_window` | Host prompt uses owned top-level alert overlay popup |
| `cmd_app_splash_live_dx_text_update` | Splash screen DxUi text update |
| `cmd_app_splash_long_run_open_close_stays_stable` | Splash screen stability |
| `cmd_app_splash_uses_dxui_surface` | Splash screen DxUi surface |
| `cmd_app_viewWidth` | View width command |

### 1.3 Shortcuts Window (31 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `cmd_app_shortcuts_column_reorder_survives_search_roundtrip` | Column reorder persists through search |
| `cmd_app_shortcuts_column_reorder_survives_sort_cycles` | Column reorder persists through sort |
| `cmd_app_shortcuts_copy_follows_reordered_columns` | Copy respects column order |
| `cmd_app_shortcuts_escape_closes_from_search_and_grid` | Escape key behavior |
| `cmd_app_shortcuts_grid_doubleClick_activates_selected_command` | Double-click activation |
| `cmd_app_shortcuts_grid_enter_activates_selected_command` | Enter key activation |
| `cmd_app_shortcuts_group_collapse_persists_through_search` | Group collapse state preservation |
| `cmd_app_shortcuts_group_collapse_persists_through_sort` | Group collapse through sort |
| `cmd_app_shortcuts_header_drag_reorders_columns_without_sort` | Header drag reorder |
| `cmd_app_shortcuts_header_resize_changes_visible_width` | Header resize |
| `cmd_app_shortcuts_header_resize_survives_search_roundtrip` | Resize survives search |
| `cmd_app_shortcuts_key_column_uses_natural_key_order` | Key column semantic sort order |
| `cmd_app_shortcuts_left_right_collapse_expand_group` | Left/Right group expand/collapse |
| `cmd_app_shortcuts_live_dx_search_interaction` | DxUi search interaction |
| `cmd_app_shortcuts_long_run_open_close_stays_stable` | Stability |
| `cmd_app_shortcuts_long_run_scrolling_stays_bounded` | Scroll bounds |
| `cmd_app_shortcuts_restores_collapsed_group_state` | Collapsed state restoration |
| `cmd_app_shortcuts_restores_combined_view_state` | Combined view state restore |
| `cmd_app_shortcuts_restores_persisted_grid_layout` | Grid layout persistence |
| `cmd_app_shortcuts_restores_persisted_sort_order` | Sort order persistence |
| `cmd_app_shortcuts_restores_reordered_sorted_grid_layout` | Reordered+sorted layout |
| `cmd_app_shortcuts_row_tooltip_tracks_hovered_cell` | Row tooltip tracking |
| `cmd_app_shortcuts_search_preserves_selection_and_group_semantics` | Search selection semantics |
| `cmd_app_shortcuts_tab_traversal_matches_expected_order` | Tab order |
| `cmd_app_shortcuts_theme_cycle_keeps_grid_legible` | Theme legibility |
| `cmd_app_shortcuts_uses_dxui_surface` | DxUi surface |
| `cmd_app_shortcuts_reordered_resized_copy_follows_visible_columns_after_search_roundtrip` | Reordered+resized copy after search |
| `cmd_app_shortcuts_reordered_resized_columns_survive_sort_cycles` | Reordered+resized layout through sort |
| `cmd_app_shortcuts_reordered_resized_copy_follows_visible_columns_after_sort_cycles` | Reordered+resized copy after sort |
| `cmd_app_shortcuts_reordered_resized_columns_survive_sort_cycles_and_search_roundtrip` | Reordered+resized layout through sort and search |
| `cmd_app_shortcuts_reordered_resized_copy_follows_visible_columns_after_sort_cycles_and_search_roundtrip` | Reordered+resized copy after sort and search |

### 1.4 Connection Manager (31 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `connection_secret_authorization_scopes` | Canonical profile/secret-kind/purpose authorization, zero/nonzero interactive timeout, unsigned tick-wrap expiry, app-run background continuation, targeted revocation, and global session clearing |
| `cmd_connection_manager_window_access_keys_focus_expected_controls` | Access key focus |
| `cmd_connection_manager_window_clean_external_reload_refreshes_list` | Clean settings hot-reload refresh |
| `cmd_connection_manager_window_close_persists_new_profile` | Close command persistence for newly created profiles |
| `cmd_connection_manager_window_dirty_external_reload_prompts_and_keeps_editing` | Dirty settings hot-reload prompt |
| `cmd_connection_manager_window_escape_from_dx_input_closes_cancel` | Escape from DxUi input |
| `cmd_connection_manager_window_enter_from_dx_input_routes_default_connect` | Enter from DxUi input routes the default Connect command |
| `cmd_connection_manager_window_applies_selected_tool_backdrop` | Shared tool-window backdrop application |
| `cmd_connection_manager_window_live_dx_interaction` | DxUi interaction |
| `cmd_connection_manager_window_masked_secret_accepts_native_chars` | Masked secret field accepts host `WM_CHAR` through the native DxUi text input |
| `cmd_connection_manager_window_textfield_doubleclick_selects_word` | Real Connection Manager text field double-click selects a word |
| `cmd_connection_manager_window_long_run_list_scrolling_stays_bounded` | Scroll bounds |
| `cmd_connection_manager_window_long_run_open_close_stays_stable` | Stability |
| `cmd_connection_manager_window_modeless_connect_posts_left_navigation` | Modeless Connect posts left-pane navigation |
| `cmd_connection_manager_window_modeless_connect_posts_right_navigation` | Modeless Connect posts right-pane navigation |
| `cmd_connection_manager_window_mtp_picker_populates_profile` | MTP editor hides auth/raw host/path fields, uses plugin-backed fake picker browse, persists the selected PnP host plus `extra.devicePuid`/`friendlyName`/`readOnly`, and keeps picker browse perf metrics present |
| `cmd_connection_manager_window_pointer_click_toggles_visible_dx_toggle` | Toggle click |
| `cmd_connection_manager_window_protocol_churn_keeps_form_and_uia_stable` | Protocol churn form/UIA stability |
| `cmd_connection_manager_window_rejects_blank_profile_name` | Blank profile-name validation |
| `cmd_connection_manager_window_rejects_duplicate_profile_name_case_insensitive` | Case-insensitive duplicate profile-name validation |
| `cmd_connection_manager_window_rejects_reserved_quick_profile_name` | Reserved Quick Connect name validation |
| `cmd_connection_manager_window_retired_dialog_files_absent` | Retired dialog file/source/project guard |
| `cmd_connection_manager_window_stale_save_prompts_before_overwrite` | Stale settings save overwrite prompt |
| `cmd_connection_manager_window_tab_traversal_live_dx_interaction` | Tab traversal |
| `cmd_connection_manager_window_theme_cycle_keeps_form_and_selection_legible` | Theme cycle form and selection legibility |
| `cmd_connection_manager_window_trims_profile_name_before_save` | Profile-name trim before save |
| `cmd_connection_manager_window_uses_dxui_command_buttons` | DxUi command buttons |
| `cmd_connection_manager_window_uses_dxui_form_action_buttons` | Form action buttons |
| `cmd_connection_manager_window_uses_dxui_form_inputs` | Form inputs |
| `cmd_connection_manager_window_uses_localized_strings_for_dynamic_labels` | Resource-backed dynamic labels |
| `cmd_connection_manager_window_wm_close_discards_new_profile` | System close/WM_CLOSE prompts before dirty discard |

Connection Manager closeout requires:

- `.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_connection_manager_window_" -TimeoutMultiplier 2`
- `.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter "cmd_connection_credential_prompt_" -TimeoutMultiplier 2`
- Debug and Release `.\build.ps1 -ProjectName RedSalamander`
- A `Specs/TestRuns/.../Commands/...` archive for the full Connection Manager window family with zero failures.

### 1.5 Connection Credential Prompt (8 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `cmd_connection_credential_prompt_access_keys_focus_expected_controls` | Access keys |
| `cmd_connection_credential_prompt_dxui_validation_and_accept` | Validation and accept |
| `cmd_connection_credential_prompt_enter_and_escape_route_default_cancel` | Enter/Escape |
| `cmd_connection_credential_prompt_escape_cancels_secret_only` | Secret-only cancel |
| `cmd_connection_credential_prompt_live_dx_interaction` | DxUi interaction |
| `cmd_connection_credential_prompt_long_run_open_close_stays_stable` | Stability |
| `cmd_connection_credential_prompt_pointer_click_toggles_secret_visibility` | Secret visibility toggle |
| `cmd_connection_credential_prompt_theme_cycle_keeps_surface_legible` | Theme cycle surface legibility and shared tool-window backdrop application |

### 1.6 Pane Commands (24+ cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `cmd_pane_changeAttributes_applies_attributes_removes_streams_and_reports` | Change Attributes selected-item scope, attribute set/clear, stream removal, and report contents |
| `cmd_pane_changeAttributes_options_dialog_uses_dxui_not_win32_template` | Change Attributes DxUi dialog, tri-state cycle back to leave unchanged, date/time rows, Include subdirectories disabled for file-only selection, UIA toggle surface |
| `cmd_pane_changeAttributes_recurse_applies_datetime_with_progress` | Recursive Change Attributes applies date/time to selected folder descendants and creates a File Operations progress task |
| `cmd_pane_changeCase` | Change case command |
| `cmd_pane_copy_text` | Copy text to clipboard |
| `cmd_pane_focusAddressBar_tab_traversal` | Address bar tab traversal |
| `cmd_pane_archive_pack_unpack_zip_roundtrip_and_validation` | Pack/Unpack ZIP round trip, sorted entries, empty-directory preservation, overwrite validation, invalid destination handling, unsupported-provider feedback, and archive perf artifact output |
| `cmd_pane_listOpenedFiles_shows_sources_prunes_closed_editors_and_focuses_items` | List Opened Files dialog empty state, viewer/editor/preview source rows, closed external-editor pruning, focus navigation, deferred close after Focus Item, and perf artifact output |
| `cmd_pane_navigationView_full_path_popup_edit_route` | Navigation full path editing |
| `cmd_pane_navigationView_full_path_popup_owned_window_activation` | Full-path popup stays open while a directly owned top-level menu window is active |
| `cmd_pane_navigationView_history_dropdown_keyboard_navigation` | History dropdown navigation |
| `cmd_pane_navigationView_path_doubleClick_enters_edit_mode` | Path double-click edit |
| `cmd_pane_navigationView_region_tab_traversal` | Navigation region tab order |
| `cmd_pane_navigationView_unfocused_pane_click_focuses_target_pane` | NavigationView actions in an unfocused pane activate that pane, navigate the clicked target, and return keyboard focus to that pane's FolderView |
| `cmd_pane_navigation_ambient_escape_returns_focus_to_active_folder_view` | Escape from main-window chrome restores active pane FolderView focus without clearing selection |
| `cmd_pane_navigation_menu_escape_returns_focus_to_active_folder_view` | Escape from the keyboard-owned pane menu restores active pane FolderView focus without clearing selection |
| `cmd_pane_navigation_nonstandard_menu_common_folders` | Nonstandard file-system NavigationView menus expose a `Common Folders` submenu with local known-folder rows and stock icons |
| `cmd_pane_navigation_status_bar_keeps_navigation_shell_stable` | Status bar focused-item detail updates, typography fit/no-clipping guard, and inactive-pane dimming |
| `cmd_pane_statusBar_uses_owned_window_and_sort_click_opens_menu` | Owned pane status-bar surface, no retained native font assignment, DxUI sort popup opening, readable popup debug state, and above-status-bar popup placement |
| `cmd_pane_navigation_directory_impact_preserves_selection` | Directory-impact refresh preserves surviving selection, drops deleted names, follows same-folder rename chains of depth 3+, and keeps renamed unselected originals unselected |
| `cmd_pane_refresh` | Refresh command |
| `cmd_pane_selection_goto_selected_name` | Go to selected name |
| `cmd_pane_selection_hide_names` | Hide selected names |
| `cmd_pane_selection_invert` | Invert selection |
| `cmd_pane_shares_shows_synthetic_rows_opens_paths_and_reports_access_denied` | Shared Directories sorted synthetic rows, open-path navigation, deferred close after Open Path, access-denied empty/error state, shortcut dispatch, and perf artifact output |

### 1.7 Find Files Dialog (61 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `cmd_pane_find_dialog_access_keys_focus_expected_fields` | Access key focus |
| `cmd_pane_find_dialog_action_buttons_activate_expected_commands` | Action button dispatch, including file Open default-open disposition without pane navigation, directory Open navigation, and Go to folder parent navigation |
| `cmd_pane_find_dialog_command_enablement_matches_idle_running_and_selection_states` | Command state management |
| `cmd_pane_find_dialog_copy_follows_reordered_columns` | Copy follows column order |
| `cmd_pane_find_dialog_directory_activation_navigates_into_selection` | Directory navigation |
| `cmd_pane_find_dialog_enter_from_checkbox_invokes_default_search` | Checkbox Enter behavior |
| `cmd_pane_find_dialog_escape_closes_popup_before_cancel` | Escape priority |
| `cmd_pane_find_dialog_escape_from_dx_control_closes_cancel` | DxUi Escape handling |
| `cmd_pane_find_dialog_exposes_live_uia_selection_and_inputs` | UIA accessibility |
| `cmd_pane_find_dialog_failure_status_is_readable` | Failure status display |
| `cmd_pane_find_dialog_grid_doubleClick_activates_selection` | Double-click activation |
| `cmd_pane_find_dialog_grid_enter_activates_selection` | Enter activation |
| `cmd_pane_find_dialog_header_click_sorts_results` | Header click sort |
| `cmd_pane_find_dialog_header_drag_reorders_columns_without_sort` | Header drag reorder |
| `cmd_pane_find_dialog_header_resize_changes_visible_width` | Header resize |
| `cmd_pane_find_dialog_large_local_search_uses_incremental_updates` | Incremental search updates |
| `cmd_pane_find_dialog_local_root_overrides_stale_context` | Local root override |
| `cmd_pane_find_dialog_long_run_open_close_stays_stable` | Stability |
| `cmd_pane_find_dialog_long_run_scrolling_stays_bounded` | Scroll bounds |
| `cmd_pane_find_dialog_mode_typeahead_updates_selection_and_dependencies` | Typeahead mode |
| `cmd_pane_find_dialog_open_parent_keeps_directory_focused_in_parent` | Open parent focus |
| `cmd_pane_find_dialog_opens_from_focused_pane_and_allows_multiple_instances` | Focused-pane initial root and multiple modeless Find windows |
| `cmd_pane_find_dialog_pointer_click_toggles_recursive_checkbox` | Recursive toggle |
| `cmd_pane_find_dialog_compact_mode_shrinks_results_grid_metrics` | Compact mode result-grid density |
| `cmd_pane_find_dialog_destination_navigation_stale_edit_host_hit_testing` | Destination NavigationView accepts delivered input after stale edit-host retirement without synthetic generation gates |
| `cmd_pane_find_dialog_editable_combo_keyboard_editing_keys` | Editable combo keyboard editing keys |
| `cmd_pane_find_dialog_recursive_local_search_and_index_availability` | Local recursive subfolder results, Path column subfolder text, shell icon indices, forced scan when indexed preference is unchecked, and indexed-backend availability |
| `cmd_pane_find_dialog_reordered_columns_survive_search_rerun` | Column order persistence |
| `cmd_pane_find_dialog_reordered_columns_survive_sort_cycles` | Column order through sort |
| `cmd_pane_find_dialog_reordered_resized_columns_survive_search_rerun` | Reorder+resize persistence |
| `cmd_pane_find_dialog_reordered_resized_columns_survive_sort_cycles` | Reorder+resize through sort |
| `cmd_pane_find_dialog_result_shortcuts_use_shell_clipboard_and_file_actions` | Result-grid configured command shortcut resolution for shell clipboard, multi-subfolder `Ctrl+X` file-drop/text fallback, two-section result context-menu clicked-item/selection dispatch, menu-driven copy-to-destination and move-to-recycle actions, compact result-action help button including captured owner/anchor backdrop, applied scrim opacity, visible-owner-resize backdrop refresh, and visible first-frame modal paint without mouse movement, action-row status plus bottom destination NavigationView embedded presentation/history seeding/live history-arrow popup behavior, stale footer NavigationView hover/click rejection after the live pointer leaves the child control, rendered first-frame destination popup state, footer-top anchored compact scroll-capped embedded destination dropdowns, explicit F5/F6 destination override, viewer/editor dispatch without scheduled Find foreground/focus restoration, row removal after accepted move/delete, canceled permanent-delete row preservation, and permanent-delete confirmation |
| `cmd_pane_find_dialog_resized_columns_survive_search_rerun` | Resize persistence |
| `cmd_pane_find_dialog_resized_columns_survive_sort_cycles` | Resize through sort |
| `cmd_pane_find_dialog_restored_combined_view_state_*` | (10 sub-cases) Combined view state restoration |
| `cmd_pane_find_dialog_restores_combined_view_state` | Combined view state |
| `cmd_pane_find_dialog_restores_persisted_grid_layout` | Grid layout persistence |
| `cmd_pane_find_dialog_restores_persisted_sort_order` | Sort order persistence |
| `cmd_pane_find_dialog_restores_reordered_grid_layout` | Reordered layout |
| `cmd_pane_find_dialog_restores_reordered_sorted_grid_layout` | Reordered+sorted layout |
| `cmd_pane_find_dialog_restores_resized_grid_layout` | Resized layout |
| `cmd_pane_find_dialog_result_drains_respect_child_input_queue_order` | Result action drains respect child input queue order |
| `cmd_pane_find_dialog_running_status_shows_phase_and_path` | Running status display |
| `cmd_pane_find_dialog_search_ops` | Search operations |
| `cmd_pane_find_dialog_service_status_shows_backend_diagnostics` | Service status |
| `cmd_pane_find_dialog_service_unavailable_warning_is_distinct` | Service warning |
| `cmd_pane_find_dialog_tab_traversal_matches_expected_order` | Tab order |
| `cmd_pane_find_dialog_theme_cycle_keeps_grid_legible` | Theme legibility plus Find results-grid folder-view visual mode and Rainbow selected-row color derived from the same containing-folder/display-name stable hash used by FolderView |
| `cmd_pane_find_dialog_uses_dxui_host_without_visible_child_controls` | DxUi host |

### 1.8 FileOps Issues Pane, Popup, and Speed Limit Prompt (24 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `cmd_pane_fileops_issues_pane_copy_follows_reordered_columns` | Copy follows column order |
| `cmd_pane_fileops_issues_pane_diff_refresh_preserves_selection` | Selection preservation |
| `cmd_pane_fileops_issues_pane_exposes_live_uia_selection` | UIA accessibility |
| `cmd_pane_fileops_issues_pane_header_click_sorts_results` | Header sort |
| `cmd_pane_fileops_issues_pane_header_drag_reorders_columns_without_sort` | Column reorder |
| `cmd_pane_fileops_issues_pane_header_resize_changes_visible_width` | Column resize |
| `cmd_pane_fileops_issues_pane_long_run_open_close_stays_stable` | Stability |
| `cmd_pane_fileops_issues_pane_long_run_scrolling_stays_bounded` | Scroll bounds |
| `cmd_pane_fileops_issues_pane_reordered_columns_survive_sort_cycles` | Column order through sort |
| `cmd_pane_fileops_issues_pane_reordered_copy_follows_visible_columns_after_sort_cycles` | Reordered copy after sort |
| `cmd_pane_fileops_issues_pane_reordered_resized_columns_survive_sort_cycles` | Reorder+resize through sort |
| `cmd_pane_fileops_issues_pane_reordered_resized_copy_follows_visible_columns_after_sort_cycles` | Reorder+resize copy after sort |
| `cmd_pane_fileops_issues_pane_restores_combined_view_state_after_recreate` | Combined view state restore |
| `cmd_pane_fileops_issues_pane_resized_columns_survive_sort_cycles` | Resize through sort |
| `cmd_pane_fileops_issues_pane_tab_keeps_grid_focus` | Tab focus |
| `cmd_pane_fileops_issues_pane_hide_restores_folder_focus` | Hiding the issues pane restores folder focus |
| `cmd_pane_fileops_issues_pane_theme_cycle_keeps_grid_legible` | Theme legibility |
| `cmd_pane_fileops_popup_global_summary_ignores_finished_tasks` | Live aggregate summary, throughput, and ETA exclude completed tasks |
| `cmd_pane_fileops_popup_progress_contracts` | Aggregate and compact progress preserve determinate versus indeterminate state |
| `cmd_pane_fileops_conflict_metadata_uses_single_provider_roundtrip` | Conflict metadata uses one basic-information provider round trip with attribute fallback only when needed |
| `cmd_pane_fileops_conflict_prompt_metadata_and_actions` | Popup conflict metadata, compact action layout, and localized action state |
| `cmd_pane_fileops_popup_presentation_settings_and_taskbar` | Popup presentation settings, minimized taskbar timer updates, and persistence |
| `cmd_pane_fileops_completed_group_and_navigation` | Completed grouping, destination navigation, and provider-qualified actions |
| `cmd_pane_fileops_issues_pane_uses_dxui_host_without_visible_child_controls` | DxUi host and shared tool-window backdrop application |
| `cmd_pane_fileops_speedLimit_prompt_uses_dxui_surface` | Speed-limit prompt DxUi surface, progress popup shared tool-window backdrop application, and file-operations popup caption glyph DirectWrite guard |
| `cmd_pane_fileops_speedLimit_prompt_live_dx_interaction` | Speed-limit prompt live input |
| `cmd_pane_fileops_speedLimit_prompt_long_run_open_close_stays_stable` | Speed-limit prompt stability |
| `cmd_pane_fileops_speedLimit_prompt_keeps_navigation_shell_stable` | Speed-limit prompt navigation-shell stability |

### 1.9 Preferences Dialog (~100 cases)

Covers every preference page with DxUi interaction, tab traversal, roundtrip restore,
and accessibility tests. Pages: General, Panes, Viewers, Editors, Mouse, Keyboard,
Themes, Plugins, Compare Directories, Hot Paths, File Operations, Advanced.

Key coverage patterns per page:
- `*_live_dx_interaction` — DxUi surface responds to input
- `*_tab_traversal_live_dx_interaction` — Tab order correct
- `*_roundtrip_restores_dxui_surface` — Settings persist through dialog reopen
- `*_page_uses_dxui_*` — Correct DxUi control types used
- `cmd_preferences_dialog_general_window_backdrop_apply_updates_supported_windows` — Preferences HWND applies the selected shared tool-window backdrop without adding a main-window acceptance assertion
- `cmd_preferences_dialog_general_page_uses_dxui_toggle_cards` — General page uses DxUi toggle/card hosts and guards HFONT-free layout through `generalUsesDxUiTypographyContext` and `generalUsesDxUiTypographyMetrics`
- `cmd_preferences_dialog_panes_page_uses_dxui_statics_and_toggles` — Panes page uses DxUi cards/toggles/combos and guards HFONT-free layout through `panesUsesDxUiTypographyContext` and `panesUsesDxUiTypographyMetrics`
- `cmd_preferences_dialog_viewers_page_uses_dxui_combo_and_button_chrome` — Viewers page uses DxUi combo/button/grid chrome and guards HFONT-free layout through `viewersUsesDxUiTypographyContext` and `viewersUsesDxUiTypographyMetrics`
- `cmd_preferences_dialog_opens_with_french_satellite_resources` — Preferences opens from the `fr-FR` satellite dialog template while still creating executable-owned custom dialog child windows
- `*_long_run_*_stays_bounded` — No resource leaks under repeated use
- `*_page_exposes_live_uia_*` — UIAutomation accessibility
- `*_search_*` — Search/filter functionality
- `*_pointer_click_*` — Mouse interaction

### 1.10 Plugin Configuration Dialog And Settings (13 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `cmd_plugin_configuration_dialog_access_keys_route_expected_actions` | Access keys |
| `cmd_plugin_configuration_dialog_enter_and_escape_route_default_cancel` | Enter/Escape |
| `cmd_plugin_configuration_dialog_live_dx_interaction` | DxUi interaction |
| `cmd_plugin_configuration_dialog_long_run_open_close_stays_stable` | Stability |
| `cmd_plugin_configuration_dialog_long_run_scrolling_keeps_dx_surface_stable` | Scroll stability |
| `cmd_plugin_configuration_dialog_pointer_click_toggles_visible_dx_toggle` | Toggle click |
| `cmd_plugin_configuration_dialog_tab_traversal_live_dx_interaction` | Tab traversal |
| `cmd_plugin_configuration_dialog_uses_dxui_command_buttons` | DxUi buttons |
| `cmd_plugin_configuration_dialog_uses_dxui_form_surface` | DxUi form |
| `settings_file_system_plugin_roundtrip` | File-system plugin settings roundtrip |
| `settings_viewer_text_plugin_roundtrip` | ViewerText plugin settings roundtrip, including rejection of transient diff-UI persistence keys |
| `viewer_text_diff_perf` | ViewerText diff perf baseline including theme-driven semantic row paint, clickable hidden-banner reveal, built-in rainbow-mode theme-switch repaint, parsed hunk-jump latency, bounded viewport growth, unresolved placeholder bands, and backtrack cache reuse |
| `viewer_text_hex_byte_color_perf` | ViewerText hex byte color perf baseline |

### 1.11 Compare Directories Options And Progress (15 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `cmd_compare_directories_create_failure_does_not_double_delete` | Injected compare-window `WM_CREATE` failure leaves no live compare window, keeps the main window alive, and allows a subsequent normal compare-window open |
| `cmd_compare_directories_window_uses_dxui_menu_bar_and_banner_buttons` | DxUi menu bar, banner buttons, banner title/progress text, no visible legacy banner text, no native font state |
| `cmd_compare_directories_leave_scope_prompt_defers_out_of_navigation_callback` | Leave-scope navigation immediately reverts during a pending run and posts the prompt after the navigation callback returns |
| `cmd_compare_directories_non_file_plugin_path_form_selection_and_empty_state` | Non-file-system plugin path-form roots, difference selection restore/invert, one-shot invert reset after pane refresh, and no-differences empty state |
| `cmd_compare_directories_options_access_keys_focus_expected_controls` | Access keys |
| `cmd_compare_directories_options_enter_and_escape_route_default_cancel` | Enter/Escape |
| `cmd_compare_directories_options_hot_reload_visible_only_and_reopen_clean` | Settings hot-reload participant lifecycle, hidden-panel silent draft reload, and stale edit cleanup on reopen |
| `cmd_compare_directories_options_live_dx_body_interaction` | DxUi body toggles/edits update and persist through the Options draft without hidden native body controls |
| `cmd_compare_directories_options_long_run_open_close_stays_stable` | Stability |
| `cmd_compare_directories_options_pointer_click_toggles_live_dx_interaction` | Toggle click |
| `cmd_compare_directories_options_scroll_to_lower_cards_stays_stable` | Scroll stability |
| `cmd_compare_directories_options_tab_traversal_live_dx_interaction` | Tab traversal |
| `cmd_compare_directories_options_theme_cycle_keeps_surface_legible` | Theme legibility |
| `cmd_compare_directories_options_uses_dxui_labels_without_visible_legacy_statics` | DxUi labels, zero visible native body/footer controls, no hidden native option state owners, and DirectWrite options typography metrics |
| `cmd_compare_directories_progress_perf` | Compare progress correctness/stability; future perf-gate use requires a self-test-local metric |

### 1.12 Settings and Infrastructure (28+ cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `folderView_empty_folder_state` | Empty folder centered state plus row-sized focused `Go to parent` placeholder item |
| `folderView_filter_watermark_empty_state` | Filter watermark display |
| `folderView_refresh_to_paint_metric_clears_after_failed_render` | Pending input-to-paint and ready same-folder refresh-to-paint metrics are both discarded after failed/no-present render paths so later unrelated successful Presents cannot emit stale latency |
| `folderView_column_widths_audit` | FolderView variable-column display and scroll audit across adversarial folder shapes, writing archived before/after metrics |
| `folderView_visible_column_widths` | Real FolderView pane verifies each column width is computed from items assigned to that visible column across Brief, Detailed, ExtraDetailed, and Thumbnails modes |
| `folderView_thumbnail_settings_roundtrip` | Per-pane thumbnail size persistence, independent left/right values, and missing-setting default of `64 DIP` |
| `folderView_thumbnail_valid_images_shell_fail` | Valid image files still display thumbnails when shell thumbnail extraction fails and WIC fallback is required |
| `folderView_thumbnail_aspect_ratio` | Non-square thumbnails preserve source aspect ratio inside the thumbnail slot |
| `folderView_thumbnail_bad_files_fallback` | Bad image-looking files complete as icon fallback, report decode failures, and leave no pending thumbnail work |
| `folderView_thumbnail_scroll_stress` | Mixed 640-item thumbnail folder keeps visible work bounded and writes scroll-stress perf evidence |
| `folderView_thumbnail_scroll_requeues_visible` | Horizontal scrolling requeues thumbnail work for newly visible columns instead of leaving valid images on icon fallback |
| `folderView_thumbnail_resize_requeues_visible` | Resizing a thumbnail pane requeues work for newly visible columns instead of leaving valid images on icon fallback |
| `folderView_thumbnail_size_change_while_pending` | Thumbnail size changes cancel stale visible work, requeue at the selected size, and settle |
| `folderView_thumbnail_stale_bitmap_messages_account_current_batch_pending` | A stale-generation payload from the current thumbnail batch decrements its owned pending apply; stale-batch and explicitly unaccounted late-current payloads do not decrement the current batch |
| `folderView_thumbnail_size_change_regenerates_fallback_icons` | Thumbnail size changes clear stale fallback icon bitmaps, requeue icon loading, and redraw fallback icons at the new target size |
| `folderView_thumbnail_return_to_normal_icon_size` | Returning from thumbnail mode to normal view uses the normal shell image-list size instead of reusing thumbnail-mode jumbo icon bitmaps |
| `folderView_thumbnail_sort_popup_slider` | Bottom-right sort popup exposes right-aligned sort shortcuts plus the four-stop thumbnail-size selector; DxUi menu coverage additionally validates progressively sized graphical stops, direct click, held-pointer drag preview, animated movement, live size text, and commit-on-release |
| `folderView_perf_large_folder_baseline` | Legacy small local FolderView baseline that verifies stock icon bitmap loading and warm local-folder smoke coverage; use `folderView_perf_huge_folder_scale` for representative large-folder evidence |
| `folderView_perf_sort_toggle_stress` | 5,000-entry adversarial folder repeatedly toggles Name, Extension, Time, Size, and None sort modes, records per-sort durations, guards inactive quick search with `incrementalSearchEffectUpdates == 0`, and emits `folder.sort_toggle_us`; this is a metric recorder, not a wall-clock threshold gate |
| `folderView_perf_scroll_render_stress` | 1,600-item normal-mode folder drives real horizontal and vertical scroll messages across Brief, Detailed, and Extra Detailed modes, repeating the gesture pass until `folder.frame.total_us` and `folder.frame.present_us` each reach the 200-sample p95 floor or the case fails; archives `metricQuality` (`folderFrameTotal`, `folderFramePresent`, `samplesEnoughForP95`, `samplesEnoughForP99`, `buildConfiguration`), records visible work, `folder.scroll_*`, and `folder.scroll.product_paint_*` metrics, asserts repeated no-op boundary scrolls do not repaint, records whether each step changed the viewport, and asserts presence of the `dwrite.text_layout.*` creation metric family (`create_count`, `create_us`, `frame_create_count`, `frame_create_us`) and the `folder.layout.*_us` phase-decomposition family (`setup`, `estimate_metrics`, `column_resolve`, `bounds`, `update_text_layouts`) |
| `folderView_perf_overlay_invalidation_stress` | Incremental-search overlay plus busy/cancel overlay animation runs for a bounded cadence window, drives deterministic warm FolderView rendering so hidden/CI launches still produce frame rows, requires overlay animation frames, requires `folder.frame.total_us` and `folder.frame.present_us` to reach the 200-sample p95 floor, and archives `metricQuality` with animation-cadence-bounded sample mode |
| `folderView_draw_item_brush_reuse_guard` | Populated FolderView selected/hovered redraw must keep the real solid-brush factory seam's draw-item transient-creation counter at zero; the committed product has no force-allocation hook inside `DrawItem` |
| `folderView_perf_huge_folder_scale` | Synthetic dummy-provider scale coverage uses 10,000 items by default and 50,000 items when `REDSALAMANDER_FOLDERVIEW_HUGE_PERF=1`; emits enumeration, first visible paint, sort toggle, quick-search keystroke-to-paint, working-set/private-byte, and bytes-per-item metrics, then select-all scrolls while asserting the live draw-item transient-brush counter remains zero |
| `folderView_perf_cold_first_visit` | Local unique-extension first-visit fixture clears the application `IconCache`, records enumeration, first paint, icon-index lookup count, icon bitmap queue count, icon-settle timing, forced first thumbnail fallback timing, and archives the OS shell-cache caveat |
| `folderView_iconcache_live_path_failure_uses_bounded_backoff` | Forces consecutive live path icon lookup failures, verifies the first and doubled negative-backoff windows suppress repeated shell calls, verifies retry after each window, and verifies success resets the negative state into a positive cache hit |
| `folderView_perf_slow_virtual_provider` | Dummy-provider latency fixture records slow enumeration and first paint, injects icon-extraction and provider-allowed thumbnail delays, proves immediate repeated failed live-path icon lookups are reissued instead of served from a negative cache via `icons.repeated_failed_lookup_count == 1`, and links paste-shortcut save latency to the ShellCommands async paste-shortcut cases |
| `folderView_perf_relayout_churn_while_scrolled` | Synthetic 10,000-item scrolled-pane relayout fixture alternates DPI, size, light/dark/high-contrast theme triggers, emits `folder.relayout_to_paint_us`, archives repaint-burst sample quality and `fontRelayoutCovered=false`, records the environment matrix, and can be rerun with `REDSALAMANDER_FOLDERVIEW_FORCE_WARP=1` for explicit WARP coverage; asserts focus and non-zero scroll position survive relayout |
| `folderView_perf_directory_change_storm` | Pane-visible local folder receives deterministic create/rename/delete/directory churn, then verifies final visible count, focus stability, and directory-change storm metrics |
| `folderView_perf_iconcache_contention` | Dual-pane icon-heavy folders with repeated unique extensions drive IconCache lock diagnostics and archive lock wait/hold evidence before any contention optimization |
| `file_action_resolution_v16_action_ids_are_case_insensitive` | File-action resolver matches action IDs case-insensitively, preserves action-definition casing, and collapses case-only references |
| `help_menu_links_external_documentation` | Help menu external documentation command placement and registry binding |
| `icon_bitmap_alpha_normalization` | Icon alpha normalization including premultiply and AND-mask transparency semantics |
| `mask_syntax_wildcards` | Wildcard mask syntax parsing |
| `menu_copy_text_group_contract` | Menu copy text contract |
| `menu_load_selection_links_restore` | Selection links restoration |
| `modeless_window_ownership` | Modeless window lifecycle |
| `navigation_location_edit_input_expands_environment_variables` | Environment variable expansion |
| `red_salamander_help_lists_diagnostics_options` | Help text documents Release diagnostics ETW and perf JSONL switches |
| `registry_integrity` | Registry settings integrity; every command has a non-empty short function-bar label with guarded examples (`MakeDir`, `UsrMenu`, `ByTime`) |
| `resource_hresult_details_format_is_valid` | Localized HRESULT details resource uses valid positional `std::format` placeholders |
| `resource_invalid_format_string_returns_raw_fallback` | Runtime resource formatting logs the failing resource ID/detail and degrades to the raw localized resource text instead of throwing or returning blank text when a localized string has invalid `std::format` syntax |
| `resource_format_placeholders_are_positional` | Product `.rc` resources reject bare `{}` and unindexed `std::format` specs while allowing documented literal file-action macros |
| `search_local_index_stream_stop_after_first` | Search stream stop semantics across a deliberately extended-length snapshot root, including long-path-safe directory creation, sibling-temp atomic save, final snapshot reload, corruption/rebuild, deletion, and first-candidate callback termination |
| `settings_file_operations_precalc_roundtrip` | Pre-calc settings roundtrip |
| `settings_file_system_plugin_roundtrip` | Plugin settings roundtrip |
| `settings_hot_reload_*` | (5 cases) Hot reload merge/suppression and transient watcher-arm self-healing |
| `settings_shortcuts_*` | (4 cases) Shortcut settings roundtrip, malformed-section rejection, and explicit unassigned sentinel persistence |
| `settings_store_search_roundtrip` | Search settings roundtrip |
| `pane_view_options_toggle_preview_pane_tabs_and_selection` | Preview pane tab-strip pointer clicks switch Folder/Preview, selected/hovered Preview close-glyph visibility, delayed Folder tab path tooltip, Preview close glyph closes preview mode, old embedded text content is cleared before rendering the next focused item, and source-pane focus is preserved |
| `pane_view_options_preview_uses_configured_embedded_viewer_and_preserves_focus` | Preview uses configured embedded viewers, keeps the same embedded instance and HWND across same-plugin image/media focus changes including `.mp4` to `.m4a`, keeps media-to-media and media-to-image switches responsive while a scope-exit-protected worker release gate deterministically retains the retiring VLC root, proves that hidden root can coexist with exactly one effectively visible and own-`WS_VISIBLE` tracked active child without a z-order traversal loop, requires zero own-style-visible children on close and one on reopen, verifies VLC child-window parenting after video-to-video and video-to-audio navigation, forwards wheel seek from VLC child surfaces, preserves source-pane focus, and persists VLC preview volume/mute state |
| `pane_view_options_preview_rejects_*_new_embedded_children` | Two focused fault-hook cases force zero detected or multiple physical new direct children, require the host to reject cardinality, hide every newly created root before plugin `Close()`, discard the viewer instance, fall back to Properties, and preserve source-pane focus |
| `pane_view_options_preview_rejects_same_plugin_*_embedded_child` | Two same-plugin reuse fault cases add a second root or invalidate the ownership-marked root and present a replacement, require the host to mark the pre-refresh child set, reject and hide the unowned new root before `Close()`, recover through a fresh validated viewer, leave exactly one effectively and own-style-visible child, and preserve source-pane focus |
| `pane_view_options_preview_uses_builtin_embedded_viewer_with_empty_associations` | Preview consults built-in embedded viewer defaults when saved viewer associations are empty before falling back to Properties text |
| `pane_view_options_preview_falls_back_to_item_properties_when_no_embedded_preview_matches` | Preview shows normalized file/folder Properties text when no specific embedded preview viewer matches, without retaining an embedded viewer instance or stealing source-pane focus |
| `pane_view_options_preview_properties_card_scrolls_and_uses_rainbow_theme` | Default no-embedded preview renders focused item Properties as DxUi cards, exposes a ScrollPanel for long metadata, accepts wheel scrolling with an increased preview scroll offset without stealing source-pane focus, and applies Rainbow theme section accents. Its named-stream coverage creates a base path longer than `MAX_PATH`, queries `FILE_NAMED_STREAMS` from that base-file handle before any capability skip, and then requires extended-length ADS creation to succeed; path-construction failures such as `ERROR_PATH_NOT_FOUND` are test failures, not capability skips. |
| `pane_filter_bar_inline_workflow` | Pane filter bar exposes an editable history combo without a redundant static Filter label, exposes the Use Filter toggle, matches the shared `selectionMasks.filterHistory` entries exactly, applies typed masks live without opening the dropdown, preserves text when toggled off, and re-applies the stored mask when toggled back on |
| `embedded_viewer_context_menus_expose_menu_actions` | Embedded menu-bearing viewers load localized menu resources for Preview right-click context menus, route selected commands through existing handlers, omit standalone-only actions and shortcut labels, trim empty groups, and do not rely on a visible embedded menubar |
| `embedded_vlc_audio_preview_stays_inside_preview` | Embedded ViewerVLC audio previews apply audio visualization as an audio-file media option, not a global VLC instance argument, so video previews do not get an extra visualizer vout and video-to-audio Preview transitions keep stable embedded playback |
| `shortcut_defaults_mapping` | Default shortcut mappings |
| `shortcut_defaults_restore_missing_*` | Startup restoration of missing default shortcut chords while preserving explicit unassigned sentinels |
| `shortcut_functionbar_dispatch_refresh` | Function bar dispatch and refresh using command short labels |

---

## 2. CompareDirectories Suite (`--compare-selftest`)

**Source:** `RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.cpp` coordinator + 4 included case files (256 runner-listed cases; 249 static `SelfTest::RunCase` call sites)

Tests the Compare Directories engine, search backends, SQLite index store,
crash quarantine, OAuth, and remote storage comparisons.

### 2.1 Core Compare Engine (37 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `accessors` | Engine accessor methods |
| `attributes` | File attribute comparison |
| `baseInterfaces` | Base interface contract |
| `cancel_completes_bounded` | Cancellation bounded completion |
| `content` | Content comparison |
| `content short reads` | Short-read content handling |
| `content_dual_io` | Dual I/O content paths |
| `content_inflight_stamp_guards_restart` | In-flight stamp restart guard |
| `content_no_io_disables_compareContent` | No-I/O content disable |
| `content_pending_elided` | Pending content elision |
| `content_queue_bounded_hi_lo` | Content queue bounds |
| `content_size_mismatch_no_pending` | Size mismatch handling |
| `contentCacheHit` | Content cache hit path |
| `concurrent_get_or_compute_decision` | Concurrent decision computation |
| `decision_cache_eviction_budget_pins_visible` | Cache eviction budget |
| `decision_cache_eviction_budget_pending_wide_tree` | Pending decision cache budget high-water |
| `decisionUpdatedCallback` | Decision update callback |
| `deep_tree` | Deep tree traversal |
| `dircache_not_polluted_by_compare_scan` | Directory cache isolation |
| `dummy_content` | Dummy plugin content compare |
| `empty_directories` | Empty directory handling |
| `ignore` | Ignore pattern matching |
| `ignore_direct_navigation_subtree` | Direct navigation into an ignored directory subtree returns empty decisions |
| `ignore_multiple_patterns` | Multiple ignore patterns |
| `ignore_pattern_count_cap` | Pattern count cap |
| `ignore_pattern_length_cap` | Pattern length cap |
| `ignore_wildcard_pathology_runtime_bound` | Wildcard pathology guard |
| `invalid_directory_entry_buffer` | Invalid entry buffer handling |
| `invalidate` | Invalidation |
| `invalidateForPath` | Path-specific invalidation |
| `missing folder` | Missing folder handling |
| `no_sync_deep_scan` | No-sync deep scan |
| `normalized_name_collision_preserves_same_side_entries` | Trailing dot/space normalized-name collisions on one side preserve both entries |
| `reparse` | Reparse point handling |
| `scan_inflight_stamp_guards_restart` | Scan restart guard |
| `setCompareEnabled` | Compare enable/disable |
| `setSettingsInvalidates` | Settings invalidation |
| `showIdentical` | Identical file display |
| `size` | Size comparison |
| `subdir pending` | Subdirectory pending state |
| `subdirattrs` | Subdirectory attributes |
| `subdirs` | Subdirectory comparison |
| `time` | Timestamp comparison |
| `typemismatch` | Type mismatch detection |
| `uiVersion` | UI version tracking |
| `unicode_filenames` | Unicode filename support |
| `unique` | Unique file detection |
| `zero_vs_nonzero_content` | Zero/non-zero content |
| `zeroByteContent` | Zero-byte file content |

### 2.2 Directory Size Callbacks (3 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `directory_size_local_callback_contract` | Local plugin directory size callback |
| `directory_size_dummy_callback_contract` | Dummy plugin directory size callback |
| `directory_size_7z_callback_contract` | 7z plugin directory size callback |

### 2.3 Search — Local Native (12 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `local_search_callback_contract` | Search callback contract |
| `local_search_content_literal` | Literal content search |
| `local_search_invalid_query_rejected` | Invalid query rejection |
| `local_search_name_and_content_and_semantics` | Name+content search |
| `local_search_name_wildcard_recursive` | Wildcard recursive search |
| `local_search_name_windows_filesystem_case_parity` | Case parity with Windows |
| `local_search_native_matches_host_fallback` | Native vs fallback parity |
| `local_search_native_unicode_long_path_matches_host_fallback` | Unicode long path parity |
| `local_search_qi_and_capabilities` | QI and capabilities |
| `local_search_scan_follow_symlink_loop_guard` | Symlink loop guard |
| `local_search_scan_wide_tree_parallel_walk_name_only` | Parallel walk |
| `local_search_backend_preferences_roundtrip` | Backend preferences |

### 2.4 Search — Host Fallback (5 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `host_fallback_search_7z_name_only` | 7z fallback search |
| `host_fallback_search_access_denied_warning` | Access denied handling |
| `host_fallback_search_content_degraded_without_io` | Content degradation |
| `host_fallback_search_dummy_name_only` | Dummy fallback search |
| `host_fallback_search_local_plugin_path_root` | Local plugin path root |
| `host_fallback_search_short_read_and_cancel` | Short read and cancel |

### 2.5 Search — Service and Indexed (30+ cases)

Covers the search service binary CLI, SQLite bootstrap, query/status roundtrip,
multi-client scenarios, rebuild control, cold start, stale root refresh, prefilter,
journal replay, snapshot reload, corruption rebuild, maintenance, deleted-root
rebuild purge, external SQLite rotation generation validation, and more.

### 2.6 Search Text Helpers (2 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `search_text_helpers_chunk_overlap_literal_and_regex` | Chunk overlap matching |
| `search_text_helpers_decoding_and_binary` | Text decoding and binary detection |

### 2.7 SQLite Index Store (6 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `sqlite_index_store_automatic_checkpoint_truncates_wal` | WAL checkpoint |
| `sqlite_index_store_automatic_compaction_is_bounded` | Compaction bounds |
| `sqlite_index_store_bootstrap_creates_schema` | Schema bootstrap |
| `sqlite_index_store_load_and_apply_journal_delta` | Journal delta |
| `sqlite_index_store_manual_compaction_reclaims_space` | Manual compaction |
| `sqlite_index_store_upgrade_paths` | Schema upgrade |

### 2.8 Infrastructure and Security (8 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `crash_quarantine_synthetic_marker` | Crash quarantine |
| `oauth_authmode_roundtrip` | OAuth auth mode roundtrip |
| `oauth_refresh_token_storage` | OAuth token storage |
| `plugin_path_math` | Plugin path mathematics |
| `try_make_relative_outside_root` | Path relativization |
| `windows_hello_cache` | Windows Hello cache |
| `connection_display_url` | Connection display URL |
| `google_drive_*` | (4 cases) Google Drive plugin contract |
| `onedrive_personal_cleared_client_id_requires_configuration` | OneDrive client ID |

### 2.9 Remote Smoke Tests (Conditional)

These cases skip when connection profiles or secrets are absent.
See `Specs/Testing/Testing_SelfTestRemoteCredentials.md`.

| Case Name | Coverage Area |
|-----------|---------------|
| `remote_file_s3` | S3 compare smoke |
| `remote_file_ftp` | FTP compare smoke |
| `remote_file_onedrive_personal` | OneDrive Personal compare |
| `remote_file_onedrive_business` | OneDrive Business compare |
| `remote_file_sharepoint` | SharePoint compare |
| `remote_s3_pagination` | S3 pagination |
| `remote_*_directory_size_callback_contract` | Remote directory size |
| `remote_ftp_continue_on_error_partial` | FTP error continuation |
| `remote_s3_metadata_smoke` | S3 metadata |
| `remote_s3_delete_missing` | S3 delete missing |

### 2.10 MTP/PTP File System (52 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `mtp_factory_single_mode_id_contract` | Single-mode factory identity |
| `mtp_identity_helpers_are_shared` | Shared MTP identity hash, suffix, sanitizer, and JSON helper source guard |
| `mtp_live_device_smoke` | Opt-in live WPD scratch write/read/overwrite/rename/copy/move/delete smoke, skipped before WPD by default |
| `mtp_queryinterface_matrix` | Required and rejected COM interface matrix |
| `mtp_json_return_buffers_are_bounded` | Two-slot JSON return buffers |
| `mtp_capabilities_are_instance_honest` | Per-instance capability JSON |
| `mtp_path_scheme_and_device_key_normalization` | Accepted schemes and connection-root normalization |
| `mtp_duplicate_names_require_stable_suffix` | Duplicate name suffix identity and fail-closed ambiguous mutation |
| `mtp_disconnect_mid_enumeration_surfaces_error` | Device-gone enumeration failure |
| `mtp_unload_quiet_point_no_callback_after_clear` | Navigation callback quiet point |
| `mtp_drive_info_and_disconnect_menu` | Drive labels and Disconnect command |
| `mtp_hung_device_times_out` | Streamed read watchdog and module quarantine |
| `mtp_watchdog_requests_backend_cancel` | Watchdog backend cancel request |
| `mtp_runtime_refresh_defers_when_worker_quarantined` | Plugin manager unload deferral |
| `mtp_mutating_create_directory_times_out` | Create-directory watchdog |
| `mtp_mutating_item_commands_time_out` | Copy/move/delete/rename watchdog |
| `mtp_writer_commit_times_out` | Writer commit watchdog with staged-byte ownership |
| `mtp_menu_and_directory_size_time_out` | Menu and directory-size watchdog |
| `mtp_reader_seek_contract` | Reader seek behavior |
| `mtp_reader_streams_on_read_not_open` | Reader open/GetSize do not materialize file bytes before the first Read |
| `mtp_concurrency_is_serialized` | Single-transport serialization |
| `mtp_backend_command_worker_is_reused` | Long-lived backend command worker reuse under concurrent host calls |
| `mtp_wpd_session_and_path_cache_reuse` | WPD session and path-object cache reuse |
| `mtp_property_fetch_is_batched` | Batched property fetch metrics |
| `mtp_public_writer_stages_until_commit` | Public writer staging and abort |
| `mtp_overwrite_byte_verify_level_matches_capability` | Configured overwrite verification path |
| `mtp_writer_overwrite_uses_temp_puid_swap` | Writer overwrite temp/PUID swap |
| `mtp_overwrite_temp_upload_failure_keeps_original_and_allows_retry` | Temp upload failure preserves original and retries |
| `mtp_overwrite_empty_temp_puid_keeps_original_and_blocks_later_upload` | PUID-less temp fail-safe and policy block |
| `mtp_overwrite_delete_original_failure_keeps_original_and_allows_retry` | Delete-original failure cleanup and retry |
| `mtp_overwrite_journal_write_failure_aborts_before_upload` | Journal-write preflight fails closed before upload/copy/move |
| `mtp_overwrite_journal_recovers_rename_temp_failure` | Retained journal replay commits temp after rename failure |
| `mtp_overwrite_journal_replay_removes_temp_when_final_exists` | Retained journal replay removes temp when final already exists |
| `mtp_overwrite_journal_replay_temp_cleanup_delete_failure_retries` | Retained journal replay retries temp cleanup delete failure |
| `mtp_overwrite_journal_recovers_committed_temp_without_tempPuid` | No-tempPUID committed-temp orphan sweep and ambiguity retention |
| `mtp_overwrite_journal_clears_completed_swap_without_temp` | No-tempPUID completed-swap inference clears final-present/temp-missing journals |
| `mtp_overwrite_journal_replay_rename_rejection_is_bounded` | Bounded retained-journal retry for rejected recovery rename |
| `mtp_overwrite_never_duplicates_or_halfwrites` | Aggregate overwrite crash-window safety matrix |
| `mtp_overwrite_verify_input_by_source_kind` | Device-source overwrite verification semantics |
| `mtp_copy_move_overwrite_uses_temp_puid_swap` | Copy/move overwrite temp/PUID swap |
| `mtp_rename_overwrite_uses_temp_puid_swap` | Rename overwrite temp/PUID swap |
| `mtp_copy_move_overwrite_temp_copy_failure_keeps_original_and_allows_retry` | Copy/move temp-copy failure preserves source/destination and retries |
| `mtp_move_fallback_delete_source_failure_leaves_duplicate_and_reports_partial` | Move fallback source-delete failure |
| `mtp_transfer_cancel_is_prompt` | Pre-transfer cancellation |
| `mtp_batch_callbacks_report_item_indices` | Batch item-completion callbacks preserve per-item indices |
| `mtp_copy_from_device_accounting` | Device read byte accounting |
| `mtp_copy_to_device_accounting` | Copy byte accounting |
| `mtp_fake_backend_enumerate_read_and_capabilities` | Fake backend browse/read/capability baseline |
| `mtp_fake_backend_move_rejects_directory_transfer_fallback` | Fake backend mirrors WPD directory move rejection |
| `mtp_fake_backend_mutations_roundtrip` | Fake backend mutation roundtrip |
| `mtp_fake_backend_readonly_configuration_blocks_mutations` | Read-only fake backend mutation block |
| `mtp_fake_backend_injection_uses_isolated_instance` | Self-test fake backend isolation |

---

## 3. FileOperations Suite (`--fileops-selftest`)

**Source:** `RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.SelfTest.cpp` coordinator + included phase files (123 runner-listed phases: 121 active phases plus setup and cleanup)

Tests file operations (copy, move, delete, rename) using a tick-driven async state machine.
Each phase represents a test case that exercises one aspect of the file operations pipeline.

### 3.1 Phase 5 — Pre-Calculation and Queue (7 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `Phase5_PreCalcSettingsApplied` | Pre-calc settings application |
| `Phase5_PreCalcCancelReleasesSlot` | Cancel releases pre-calc slot; preflight copy card exposes collapse, Skip, Speed Limit, and Cancel |
| `Phase5_PreCalcCancelLatencyLocal` | Cancel latency measurement |
| `Phase5_PreCalcSkipContinues` | Skip continues operation |
| `Phase5_CancelQueuedTask` | Queued task cancellation |
| `Phase5_SwitchParallelToWaitDuringPreCalc` | Parallel to wait mode switch |
| `Phase5_SwitchWaitToParallelResume` | Wait to parallel resume |

### 3.2 Phase 6 — Popup and Bandwidth (5 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `Phase6_PopupRateSmoothing` | Popup rate/ETA smoothing contract |
| `Phase6_PopupSmokeResizeAndPause` | Popup resize and pause UI |
| `Phase6_DeleteBytesMeaningful` | Delete bytes tracking |
| `Phase6_LocalBandwidthThrottle` | Local bandwidth throttling |
| `Phase6_ParallelBandwidthThrottleFairness` | Parallel bandwidth fairness |

### 3.3 Phase 7 — Watcher, Concurrency, and Batching (16 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `Phase7_WatcherChurn` | Directory watcher churn |
| `Phase7_CacheBorrowNoWatchInvalidation` | Cache borrow without watch invalidation |
| `Phase7_CrossPaneVisibleRefreshLocal` | Cross-pane refresh (local) |
| `Phase7_CrossPaneVisibleRefreshDummy` | Cross-pane refresh (dummy) |
| `Phase7_CrossPaneRelocateLocal` | Cross-pane relocate |
| `Phase7_LargeDirectoryEnumeration` | Large directory enumeration |
| `Phase7_ParallelCopyMoveKnobs` | Parallel copy/move knobs |
| `Phase7_CopyMoveConcurrency16Perf` | 16-thread concurrency performance |
| `Phase7_CopyRecursiveParallelismMatrix` | Recursive copy/move matrix including copied reparse items, nested concurrency 1, forced and optional real cross-volume move fallback, partial error, and active-worker cancellation |
| `Phase7_AutoConcurrencyHints` | Auto concurrency hints |
| `Phase7_PerItemDirectoryCopyInFlightLines` | Per-item directory copy in-flight |
| `Phase7_CopyItemsSingleFolderRecursiveParallelism` | Single-folder recursive per-item parallelism |
| `Phase7_CopyItemsMultiRootUnevenRecursiveParallelism` | Multi-root uneven recursive per-item parallelism |
| `Phase7_SharedPerItemScheduler` | Shared per-item scheduler |
| `Phase7_ParallelDeleteKnobs` | Parallel delete knobs |
| `Phase7_RecycleBinBatchDelete` | Recycle bin batch delete |
| `Phase7_RecycleBinBatchDeleteMultiBatch` | Multi-batch recycle bin delete |

### 3.4 Phase 8 — Defaults and Validation (5 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `Phase8_DefaultBandwidthLimitFromSettings` | Default bandwidth limit |
| `Phase8_TightDefaults_NoOverwrite` | Tight defaults without overwrite |
| `Phase8_InvalidDestinationRejected` | Invalid destination rejection |
| `Phase8_InvalidSizeBytesRejected` | Invalid size rejection |
| `Phase8_PerItemOrchestration` | Per-item orchestration |

### 3.5 Phase 9 — Conflict Prompts (8 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `Phase9_ConflictPrompt_OverwriteReplaceReadonly` | Overwrite readonly |
| `Phase9_ConflictPrompt_ApplyToAllUiCache` | Apply-to-all UI cache |
| `Phase9_ConflictPrompt_OverwriteAutoCap` | Overwrite auto cap |
| `Phase9_ConflictPrompt_LocalFileOntoDirectory` | Local file-on-directory metadata suppresses Overwrite |
| `Phase9_ConflictPrompt_SkipApplyToAll` | Skip + All similar decision cache |
| `Phase9_ConflictPrompt_RetryCap` | Retry cap |
| `Phase9_ConflictPrompt_SkipContinuesDirectoryCopy` | Skip continues directory copy |
| `Phase9_PerItemConcurrency` | Per-item concurrency |

### 3.6 Phase 10–15 — Advanced Operations (12 cases)

| Case Name | Coverage Area |
|-----------|---------------|
| `Phase10_PermanentDelete` | Permanent delete confirmation, cancellation guard, and confirmed execution |
| `Phase11_CrossFileSystemBridge` | Cross-filesystem bridge |
| `Phase11_BridgeSingleFolderParallelCopyInFlightLines` | Bridge single-folder parallel |
| `Phase11_BridgeMultiFolderParallelCopyInFlightLines` | Bridge multi-folder parallel |
| `Phase11_BridgePipelineDummyToDummyPerf` | Bridge pipeline performance |
| `Phase11_ConnectionOverridePrecedence` | Connection override precedence |
| `Phase11_ConnectionOverrideGlobalGate` | Connection override global gate |
| `Phase11_ConnectionOverrideClamp` | Connection override clamp |
| `Phase12_ReparsePointPolicy` | Reparse point policy |
| `Phase13_PostMortemDiagnostics` | Post-mortem diagnostics |
| `Phase14_PopupHostLifetimeGuard` | Popup host lifetime and reentrant visibility/placement guard |
| `Phase15_FileSystem7zReadSeekSmoke` | 7z read/seek smoke |
| `Phase15_FileSystem7zMountPathImpact` | 7z mount path impact |

### 3.7 Phase 16 — Remote Storage (Conditional, 16 cases)

These cases skip when connection profiles or secrets are absent.

| Case Name | Coverage Area |
|-----------|---------------|
| `Phase16_RemoteWatchContractExposure` | Remote watch contract |
| `Phase16_RemoteFtpSecret` | FTP secret validation |
| `Phase16_RemoteFtpSandbox` | FTP sandbox operations |
| `Phase16_RemoteSftpSecret` | SFTP secret validation |
| `Phase16_RemoteSftpSandbox` | SFTP sandbox operations |
| `Phase16_RemoteScpSecret` | SCP secret validation |
| `Phase16_RemoteScpSandbox` | SCP sandbox operations |
| `Phase16_RemoteImapSecret` | IMAP secret validation |
| `Phase16_RemoteImapSandbox` | IMAP sandbox operations |
| `Phase16_RemoteS3Secret` | S3 secret validation |
| `Phase16_RemoteS3Sandbox` | S3 sandbox operations |
| `Phase16_RemoteS3FileOps` | S3 file operations |
| `Phase16_RemoteOneDrivePersonalSecret` | OneDrive Personal secret |
| `Phase16_RemoteOneDrivePersonalSandbox` | OneDrive Personal sandbox |
| `Phase16_RemoteOneDrivePersonalFileOps` | OneDrive Personal file ops |
| `Phase16_RemoteOneDriveBusinessSecret` | OneDrive Business secret |
| `Phase16_RemoteOneDriveBusinessSandbox` | OneDrive Business sandbox |
| `Phase16_RemoteSharePointSecret` | SharePoint secret |
| `Phase16_RemoteSharePointSandbox` | SharePoint sandbox |

### 3.8 Cleanup (1 case)

| Case Name | Coverage Area |
|-----------|---------------|
| `Cleanup_RestorePluginConfig` | Restore plugin configuration after test |

---

## 4. Performance Tests (`PerformanceTests2/`)

**Framework:** Microsoft CppUnitTest (DLL)

| Test Class / File | Cases | Coverage Area |
|-------------------|-------|---------------|
| `FolderIconEnumerationPerfTest` | 1 | Icon enumeration caching under load |
| `FolderIconEnumerationDuplicatePathPerfTest` | 1 | Duplicate path icon edge cases |
| `FolderViewColumnLayoutTests` | 5 | Pure variable-column width calculation, detailed/metadata line locality, first-column scroll-stop/thumb-release snapping, leading-gutter hit-test semantics, and FolderView sort parallel-threshold policy |
| `FolderViewRefreshDuplicatePathPerfTest` | 2 | FolderView refresh with duplicate paths and compact-mode hit testing |
| `PerformanceTests2.cpp` | 3 | Splash close guard and empty plugin-manager discovery failures |

---

## 5. PoC Tests

These harnesses are currently executable-level gates, not shared per-case JSON
reporters. DxUiTests has suite filtering (`--suite=<name>`) but many assertions
fail fast via `Require(...)->std::exit(1)`. ViewerPETests,
ViewerSqliteTests, MonitorTest, and LocalizationTests use local success
aggregation. Adding true per-case machine-readable reporting requires a named
case registry/common reporter contract for these harnesses.
Run HWND focus-sensitive DxUi suites such as `NativeTextInput` serially rather
than as parallel foreground-window peers when collecting closeout evidence;
they create real test windows and can legitimately affect process/global
Win32 focus. Standalone DxUiTests that assert real Win32 focus, caret, or
foreground ownership must route through `TryFocusDxUiTestWindow` or
`TryActivateDxUiTestWindow` and must emit an explicit `SKIPPED:` reason when
the current desktop cannot satisfy that environmental precondition.

`Tools\Run-AllTests.ps1 -Suite CI` is the GitHub Actions PR gate lane. It runs
the three in-product suites as separate processes, splits DxUiTests by suite,
preserves the explicit ViewerPE prompt cases, adds FileSystemCurlTests and
RedConfigureTests, runs PluginContractTests/SettingsSchemaTests/CrashHandlingTests,
keeps monitor ETW latency in Suite Full only, and excludes only
`RequiresBuildToolchain` Pester cases from artifact-only CI. The pull-request
workflow also performs a Debug ARM64 build through the reusable build workflow
without executing ARM64 output on the x64 hosted runner. Suite CI also
enables blocking failure classification: standalone pass-on-rerun is `FLAKY`,
fail-again is `REGRESSION`, and broad in-product suite fail plus isolated case
pass now runs shuffle triage across three seeds. Pass-all-shuffle evidence is
blocking `FLAKY`, fail-any-shuffle evidence is blocking `REGRESSION`, missing
shuffle evidence remains `ISOLATION_SUSPECT`, and retry attempts preserve the
triage `shuffle_seed` in the aggregate artifact. The same runner validates
`Tools/test-quarantine.jsonl`; any invalid, expired, or active quarantine entry
remains a blocking result with owner/expiry/repair metadata in the aggregate
artifact.

`Tools\Run-AllTests.ps1 -Suite Full` must preserve stdout/stderr for standalone
EXE and CppUnitTest entries through per-suite `*.output.log` files and
`output_log_path` in `run-all-tests-results.json`; an executable-only exit code
is not enough for closeout triage. ViewerPETests has nested fresh-process
coverage: normal isolated viewer cases use the default 120-second process cap,
while the six-cycle `TestViewerShellComboHostsLongRunOpenCloseStayStable`
stress entry has its own 600-second outer cap so the parent harness does not
kill valid nested churn before the per-child checks can report their result.

`.github/workflows/nightly-flake.yml` owns the expensive scheduled/manual
shuffle-repeat lane. It reuses the Debug self-test build artifact and runs
`Tools\Run-AllTests.ps1 -Suite All -SkipBuild -SelfTestRepeat 5
-SelfTestShuffleSeed <seed> -ClassifyFailures`, uploading
`selftest-artifacts-nightly-shuffle`. This lane is deliberately separate from
the PR CI workflow.

Current DxUi native text-input coverage includes UIA `TextUnit_Line` endpoint
and selected-range movement over wrapped multiline `TextField` visual lines,
derived from native caret geometry instead of logical newlines alone.
It also covers deterministic native IME `GCS_CURSORPOS` and `GCS_COMPCLAUSE`
diagnostics, mapping cursor and clause offsets to absolute retained-text indexes.
Multiline/wrapped native IME payload coverage now includes a live preview followed
by `GCS_RESULTSTR`, proving the result replaces the original composition anchor
instead of deleting preview-length text from the preserved base state.
UIA `TextPattern` range geometry also covers multiline mixed-BiDi selected
ranges, comparing provider rectangles with the retained DirectWrite
`HitTestTextRange` geometry. Single-line mixed-BiDi selected ranges now use the
same retained DirectWrite range-rectangle hook instead of a text-viewport
fallback.
UIA Text/TextEdit event coverage includes deterministic counters for native
retained text changes, retained selection changes, IME composition text changes,
retained caret moves, and IME conversion-target changes while the production
path raises the matching UI Automation provider events, including active text
position events with a collapsed caret range.
Native TSF activation coverage verifies that focusing a native `TextField`
activates a TSF thread manager/document manager/context over the focused
`ITextStoreACP`, keeps it alive while text focus remains, and releases it when
focus leaves.
Native TSF composition-owner coverage verifies the text store exposes
`ITfContextOwnerCompositionSink`, accepts composition start callbacks, allows
composition, and accepts update/end callbacks.
Native TSF ACP2 coverage verifies the same retained native text store exposes
`ITextStoreACP2` and answers ACP2 text, selection, screen extent, text extent,
and point-to-ACP queries while under a TSF read lock.
Native TSF external-change soak coverage verifies repeated retained emoji text
changes emit one bounded sink notification set per observed change, even when
the sink requests a synchronous read lock from the text-change callback.

| Project | Cases | Coverage Area |
|---------|-------|---------------|
| **DxUiTests** | ~50+ | DxUi color parsing, theme rendering, control creation, HSL/RGB conversion, submenu cascade hover timing, single-line text selection clipping, compact TextField density default vertical-padding coverage with explicit-padding override semantics, native RTL/mixed-BiDi selection clipping outside visible clear/reveal trailing buttons, selected TextField emoji color-font rendering, native TextField selected/unselected/multiline/mixed-BiDi and editable ComboBox emoji color-font rendering without a hidden bridge child plus color-glyph pixel-count perf rows, native masked emoji color-font suppression and unmask restore, TextField/native extended emoji text-element deletion and Shift+Arrow selection for ZWJ sequences, variation selectors, skin-tone modifiers, and regional-indicator flags, native emoji copy/cut/paste selection replacement, clipboard round-trip, and undo/redo coverage for grinning face, woman technologist, rainbow flag, skin-tone modifier, and regional-indicator flag text elements, native pointer hit-test snapping over extended emoji text elements for TextField and editable ComboBox, native masked exact-policy one-dot-per-text-element state for extended emoji, native concealed-policy privacy display ranges with same-bucket edit stability plus full-reset/refocus epoch regeneration, hidden concealed pointer end-snap plus keyboard edit/paste/undo/redo coverage, secret render/display-dot/reveal-toggle perf rows, native masked reveal-button pointer and keyboard press-and-hold peek without clearing the secret, keyboard release/blur remask, Tab traversal through the reveal affordance, reveal-button UIA Button/Invoke provider coverage with masked value/text non-disclosure after Invoke, explicit `PasswordRevealMode::Hidden` no-affordance coverage, and explicit `PasswordRevealMode::Visible` persistent plaintext/copy coverage across blur/read-only/disabled transitions, native host text-input keyboard routing, native text-input backend focus/session/caret scaffolding, backend-neutral `SupportsTextInput()` consumer coverage for `TextField` and editable `ComboBox`, native editable-combo session/typing coverage without a hidden bridge child, native editable ComboBox Ctrl+A/C/X/V/Z/Y, Shift+Insert, Shift+Delete, normalized paste, Alt+Down popup-open, Escape popup-close, retained selection, and native-session plus backend-neutral `TextInputState` sync coverage, native inherited flow-direction session state and focused inherited-flow refresh, shared single-line DirectWrite reading-direction-aware visible layout/caret/hit-test/selection-paint plumbing for `TextField` and editable `ComboBox` plus TSF point/extents and UIA RangeFromPoint fallback with `dxui.textinput.bidi_hit_test_us` / `dxui.textinput.bidi_caret_rect_us` perf rows, native key-to-state and key-to-paint perf rows from a deterministic typed-and-rendered native TextField scenario, native edit-transaction and undo-depth perf rows for direct edits, undo, and redo, native no-op delete transaction suppression plus once-per-mutation text-change notifications, native pointer caret-placement state sync, native host-HWND single-line double-click, synthesized repeated-click word selection, third-click select-all, drag-selection replacement over punctuation-delimited text, mixed-BiDi drag selection across Latin/Hebrew script boundaries in both LTR and RTL visual directions with logical UTF-16 clipboard order, native mixed-BiDi pointer hit-test matrix coverage for pixel-rounded leading/middle/trailing DirectWrite visual spans in both LTR and RTL flow directions, native BiDi scenario matrix coverage for pure LTR, pure RTL, Arabic plus Latin digits, surrogate pairs inside RTL text, and path-like RTL host text, native BiDi keyboard logical-boundary coverage for Home/End, Ctrl+End, Shift+Home/End, logical Left/Right, Backspace, and Delete in an RTL host, and native mixed-BiDi edit transaction coverage for logical-order copy/cut/paste, undo/redo selection restoration, Ctrl+Backspace, and Ctrl+Delete around mixed-script word/separator boundaries, native surrogate-pair and extended emoji backspace/delete state sync, native Ctrl+Backspace/Ctrl+Delete word-deletion state sync, native root-reset teardown, native focused-field bounds-change caret refresh, native Tab/default/cancel/context-menu/WM_SYSCHAR routing including attached logical/wrapped multiline `VK_APPS` and `Shift+F10` context-menu keys through the host HWND without a hidden bridge child, native IME start/end composition-state lifecycle, no-payload IME suppression without active composition, read-only IME composition suppression, composition-over-selection range tracking, composition-owned Return/Escape/Tab routing, NavigationView edit-suggest active-composition Down-arrow ownership, host-owned IMM32 composition/candidate window placement at the native caret, moved-field, multiline/wrapped caret-line and focused-control move anchoring, editable ComboBox move, programmatic `TextField` and editable `ComboBox` caret movement, focused `TextField` padding-change and editable `ComboBox` density-change reanchoring, multiline-scroll, and DPI IME reanchoring, native IME result commit, active composition preview with retained composition/conversion-target underline paint geometry for `TextField` and editable `ComboBox`, preview-then-result commit against the original multiline/wrapped IME base anchor, cancel restore, masked UIA `IsPassword`, ValuePattern, and TextPattern non-disclosure, explicit UIA HelpText exposure from retained controls, UIA TextPattern/TextEditPattern document/selection ranges, TextRange clone/endpoint comparison, RangeFromPoint leading-edge caret mapping plus multiline native hit-test mapping, text-element-aware character-unit endpoint/range movement over ZWJ emoji clusters, word-unit endpoint/collapsed/noncollapsed range movement, logical line endpoint/selected-range movement for newline-delimited multiline `TextField` content, multiline TextField non-exposure of ValuePattern, host-thread-dispatched TextField/editable ComboBox range `Select()`, and non-empty selected-range bounding rectangles plus simple LTR same-visual-line, newline-delimited multiline caret-geometry, wrapped multiline visual-line, and single-line plus multiline mixed-BiDi DirectWrite selected-range rectangles for `TextField`, UIA TextPattern/TextEditPattern document ranges plus RangeFromPoint and retained selection for editable `ComboBox`, `dxui.uia.text_range_us` perf rows, native IME TextEdit active-composition/conversion-target ranges, direct native TSF `ITextStoreACP` / `ITextStoreACP2` lock/text/end-ACP/selection/basic geometry/point-to-ACP/mutation/mixed-BiDi text-viewport point/extents/same-line and wrapped multiline text extent/multiline and wrapped point-to-ACP mapping/SetText replacement/query-only insert metadata/layout-unavailable/store-originated and retained-external sink notification plus UnadviseSink identity, read-write edit-transaction, and reentrant-lock rejection coverage and logical UTF-16 emoji range selection/replacement for focused `TextField`, direct native TSF retained selection and insert-at-selection mutation coverage for focused editable `ComboBox`, native single-line and multiline clipboard/undo routing, native host edit-message routing including no-selection `WM_CLEAR`, native masked-hidden clipboard suppression, native masked-revealed copy/cut mutation plus remask on blur/read-only/disable, before Escape cancel, on window deactivation, and on reveal-button capture loss, native read-only mutation suppression, NavigationView native DxUi host-backed address/full-path edit routing without a bridge subclass, NavigationView invalid-path retained HelpText validation feedback, FolderView incremental-search helper behavior, inactive-pane visual-state helpers, and empty-folder placeholder layout metrics |
| **DxUiTests / NativeTextInput** | 118 | Includes native `TextField` and editable `ComboBox` active IME composition/conversion-target inline underline paint geometry derived from retained range rectangles, programmatic retained caret movement reanchoring active IMM32 composition/candidate forms, focused `TextField` padding-change and editable `ComboBox` density-change reanchoring of active IMM32 composition/candidate forms, focused read-only and masked state cache refresh while a native session is active, editable `ComboBox` active IME composition/candidate reanchoring after focused bounds changes without creating a hidden bridge child, native multiline/wrapped multiline IME composition/candidate anchoring across logical/visual caret lines and focused-control bounds changes on the host HWND, native host-HWND focus-loss native-session teardown/regain while retaining logical text focus, native multiline/wrapped Return default-button suppression plus Tab/Shift+Tab traversal and Escape cancel routing, native multiline/wrapped host-HWND character and Return replacement state sync, editable `ComboBox` exact-match selection plus delete/word-delete command sync coverage on the native host HWND, native single-line tab-character suppression plus partial-selection paste state sync, native Win32 edit-message protocol coverage for `WM_GETTEXT`, `WM_SETTEXT`, `EM_GETSEL`, `EM_SETSEL`, and `EM_REPLACESEL`, native multiline/wrapped multiline IME composition-owned Return/Escape/Tab routing, modified navigation-key routing during active IME composition, host/app deactivation teardown of active IME composition, native IME preview/result commit followed by `WM_IME_ENDCOMPOSITION`, and native IME result-only versus continuing-composition host-key routing coverage. |
| **FileSystemCurlTests** | live inventory | IMAP UIDVALIDITY-qualified leaf identity, RFC2047 subject decoding, mailbox `STATUS`/CAPABILITY parsing, stale-identity rejection, fake-mailbox safe single-message delete/rollback, constant security and Properties command-count models, listing summary repair batching, and bounded per-listing repair fetch budget coverage. |
| **PluginContractTests** | 400+ assertions | Plugin factory/enumeration/schema/capability contracts; transactional configuration rollback and unknown-member round-trip across local, 7z, Curl, Google, Microsoft, and S3 providers; legacy 7z/Curl secret import-without-repersistence; local contradictory writer-flag attribute preservation; Dummy late-collision copy/move no-partial-state and descendant read-only delete policy; configuration-gated plugin debug selftests (including cancellable/retryable 7z index construction; Microsoft Drive credential-bound Graph/upload URL validation, server-acknowledged upload offsets, concurrent-child-safe merge cleanup, committed-move cleanup debt, redacted diagnostics, and continuation rejection; Google Drive response/deadline caps, eight-caller single-flight refresh, bounded retry, repeated-token rejection, and reversible case-sensitive identity; S3 bounded paging, exact upload-byte proof, final directory-size callback propagation, recursive-delete convergence, committed-cleanup debt, retryable multipart abort, and AWS initialization state); eight alternating Curl/Google physical-unload cycles proving the survivor still creates a curl handle; a host-facing S3 runtime-refresh unload quiet-point test; and direct `PackedFileInfoBuffer` empty/multi-entry metadata, x64/ARM64 alignment, null-output, out-of-range, malformed-name-size, and malformed-offset contracts. |
| **RedConfigureTests** | 39 | RedConfigure page definitions and four-page UI smoke, workspace discovery, repo-sized scan/parse/validate performance, theme JSON5 parsing/grouped export/validation, SettingsStore parser parity, RC string/menu/dialog parsing plus SDK compiler acceptance, placeholder/unbalanced-brace validation, translation views, all seven localization batch kinds with stable-identity stale rejection, combined typed validation and undo/redo, all ten theme recipes with exact-reference/direct-source-only/invalid/stale/atomic contracts, 512-token mass-preview budget, duplicate-theme boundary/result behavior, invariant numeric grammar under a comma-decimal locale, atomic reparsed export, and BOM-less UTF-16 RC loading. |
| **DxUiTests / Accessibility** | 32 | Includes editable `ComboBox` single-line mixed-BiDi DirectWrite selected-range rectangle coverage through UIA `TextPattern::GetSelection()` / `TextRange::GetBoundingRectangles()`, preserving logical UTF-16 selected text while comparing screen rectangles against retained `ComboBox::TryGetTextInputRangeRects(...)` geometry, horizontally scrolled `DxUi::Grid` row UIA coverage proving row names and row-owned `GridCell` fragments include off-view model columns, `DxUi::ScrollPanel` UIA point-hit coverage proving scrolled-out content-space children are clipped while visible children resolve at viewport-translated points, cross-thread UIA action dispatch timeout coverage proving late host-thread handlers write only into heap-owned dispatch storage after the sender has timed out, and Label-only root parity coverage proving the snapshot path does not collapse labels into direct semantic roots. |
| **DxUiTests / ReadOnly** | 24 | Focused read-only multiline/wrapped text-field coverage, including attached native/default-host cases with no bridge opt-in host wrapper that prove host `WM_COPY` copies logical and wrapped multiline text, no-selection host `WM_COPY` is a clipboard no-op, no-selection host `WM_CUT`/`WM_CLEAR` leave clipboard/text/caret unchanged, copy shortcuts preserve full selection, no-selection copy/cut shortcuts leave clipboard/text/caret unchanged, undo/redo no-ops preserve full selection, Ctrl+Backspace/Ctrl+Delete no-ops keep native caret state stable, Ctrl+Arrow word navigation syncs native caret state, and `WM_CUT`, `WM_PASTE`, `WM_CLEAR`, and `WM_CHAR` are suppressed without creating a hidden bridge child. |
| **LocalizationTests** | ~5 | Resource owner registration, satellite string/menu/dialog lookup, localized dialog templates with executable-owned custom child classes, fallback to embedded resources, and persisted `ui.language` roundtrips |
| **ViewerPETests** | ~20+ | PE viewer plugin, image viewer, text viewer including AppTheme-driven parsed diff semantic colors with runtime theme switching, explicit rainbow-mode coverage, and high-contrast coverage, base-background unchanged rows plus dim diff-marker metadata, clickable hidden-context banners in hunks-only diff mode, non-anchor parsed hunk presentation, pane-local side-by-side visual-layout metadata and visible split-row counters for parsed diff viewports, split-row top-visible text snapshot coverage after hunk navigation, parsed hunk count and active-hunk snapshot metadata, next/previous hunk navigation, diff presentation, lazy referenced-file expansion, parsed-document reuse across diff variants, active-section-only unchanged-text hydration with on-demand section jumps, viewport-windowed unchanged-row rehydration on scroll, range-bounded referenced-file reads for viewport-nearby unchanged context, cached reuse when revisiting already hydrated viewport ranges, referenced-file content reuse across expanded layouts, hatched placeholder-gap metadata for unresolved rows, diff section navigation, horizontal scrolling in both parsed diff layouts after wrap is disabled including side-by-side to inline presentation switches, shared viewer combo-host popup expansion/collapse, compact combo chrome, Escape/Tab focus return without closing from chrome, embedded standalone-combo/chrome hiding including ViewerSpace, Image/RAW header-combo inset, and clean idle close behavior across `ViewerPE`, `ViewerWeb`, `ViewerImgRaw`, and `ViewerText`, larger fully buffered diff parsing including promotion beyond the normal text buffer size for both extension-recognized and header-sniffed diffs, raw streamed multi-file section indexing/navigation beyond the fully buffered parse cap, parser-fallback coverage, ViewerSpace window/menu hosting plus Direct2D tooltip overlay width/native-tooltip regression coverage and Escape scan-cancel/idle-close behavior, and VLC viewer |
| **ViewerSqliteTests** | ~15+ | SQLite engine, query execution, schema inspection, UI rendering, results-grid-first keyboard focus, Tab/Shift+Tab traversal order, and Escape focus-return/posted-close behavior |
| **MonitorTest** | ~6 | ETW TraceLogging provider emit/receive validation plus compile-time/runtime guards that invalid rectangle visualization remains opt-in, normal Debug and Release monitor self diagnostics suppress Info/Perf output unless the runtime ETW flag is enabled, normal monitor display rejects self-originated ETW events, and ColorTextView scrollbar visibility reaches a stable minimal state |

---

## 6. Tooling Script Tests

| Test File | Coverage Area |
|-----------|---------------|
| `Tools\Tests\BuildOutputProcess.Tests.ps1` | Build preflight self-test protection, exclusive lock-file identity and managed-thread nesting, contained direct-child delegation, parallel-runspace and stale-descendant exclusion, owner diagnostics, abandonment contamination (including stale v1 owner migration), residual compiler diagnostics, exact-path matching, and ordinary interactive-instance close behavior |
| `Tools\Tests\BuildProjectSelection.Tests.ps1` | Project selection and direct vcxproj builds |
| `Tools\Tests\BuildReproducibility.Tests.ps1` | Declarative runtime dependency staging, missing-file failure, clean portable extraction, pinned vcpkg identity, Release test definitions, bounded build parallelism, ARM64 PR compilation, and seeded ASan workflow contracts |
| `Tools\Tests\MSBuildInvocation.Tests.ps1` | MSBuild invocation planning and diagnostic parsing |
| `Tools\Tests\ModalWindowShellSourceContracts.Tests.ps1` | About/Fatal Error shared modal-shell adoption, owner restoration, canonical DxUi modal-loop reuse, preserved `GetLastError`/`WM_QUIT` behavior, owned-HWND reset, and session-end exclusion |
| `Tools\Tests\HwndRenderTargetResourcesSourceContracts.Tests.ps1` | FunctionBar/StatusBar adoption of the narrow shared D2D factory/HWND-target/solid-brush lifecycle, target-dependent reset/factory-retention boundaries, local typography ownership, teardown release, and stable paint metrics |
| `Tools\Tests\PackedFileInfoBufferSourceContracts.Tests.ps1` | Checked/aligned packed `FileInfo` owner, exact six buffered-provider facade adoption, local streaming/Dummy fixture exclusions, and removal of provider-local sizing/traversal copies |
| `Tools\Tests\PluginLifetimeConsolidationSourceContracts.Tests.ps1` | Google/Microsoft Drive registration callback-state adoption, centralized callback-return module-pin transfer, and canonical manager plugin-ID lookup/lease contracts |
| `Tools\Tests\MtpLiveCloseout.Tests.ps1` | MTP live closeout wrapper safety, archival, and environment-restoration contracts |
| `Tools\Tests\ProcessStreaming.Tests.ps1` | Process output streaming, logging, and kill-on-close descendant containment |
| `Tools\Tests\RedSalamanderPluginDeployment.Tests.ps1` | Targeted RedSalamander build repopulates sibling binaries/plugins and RedSalamander dependency-linked plugin language resources; launches the child build through the sanitized environment helper to canonicalize duplicate `Path` / `PATH`; tagged `RequiresBuildToolchain`, excluded from artifact-only CI test jobs, and bounded with captured build logs |
| `Tools\Tests\ResourceLocalizationContracts.Tests.ps1` | Resource placeholder positional-order and satellite placeholder-equivalence contract |
| `Tools\Tests\RunAllTestsPlan.Tests.ps1` | CI/Full runner test-plan enumeration, unified test-sandbox root selection, Pester 3/newer invocation compatibility, dead-PID stale run cleanup, disk-audit evidence, and legacy cleanup target planning including locked-target failure reporting, optional native perf-budget gate forwarding, repeat/shuffle and injected classifier-proof argument forwarding, focused-filter shuffle-triage planning, invariant timeout formatting, blocking failure classification with shuffle-triage evidence, reviewed quarantine validation and repair-lane planning, GitHub step summary formatting, aggregate artifact, per-case history/dashboard output, and result-coverage validation |
| `Tools\Tests\SanitizedEnvironment.Tests.ps1` | Child process environment normalization |
| `Tools\Tests\TestHarnessSourceContracts.Tests.ps1` | Source guards for test harness CLI/error handling, case-listing, repeat/shuffle controls, injected classifier-proof hooks, native unified TestSandbox root/run-id consumption, native TestSandbox scratch acquisition, FileOperations alternate-volume TestSandbox scratch routing/pruning, RedConfigureTests, ViewerSqliteTests, CrashHandlingTests unified TestSandbox scratch routing, Commands plugin-config native TestSandbox routing, ShellCommands shortcut-save native TestSandbox routing, PerformanceTests2 unified TestSandbox scratch routing, Commands BatchRename window fixture native TestSandbox routing, ViewerPETests unified TestSandbox fixture routing, Compare foreground service sandboxed stdout capture and kill-on-close JobObject isolation, broad raw temp/profile/legacy-root source guard, local index snapshot reload advisory timing, raw self-test wait scaling through `SelfTest::ScaleTimeout`, legacy TestSandbox cleanup script safety, failed-status warning behavior, and runner invocation, CompareDirectories and FileOperations explicit-order seeded shuffle, FileOperations native repeat aggregation, result-emission, duplicate-name contracts, CompareDirectories listed-case coverage, self-test fatal-modal bypass, partial crash-result preservation, file-operations prefix filters, FileOperations pause-point centralization, FileOperations bridge IO decorator source-size fault seam, FileOperations bridge create-directory race env-helper reuse, FileOperations issues-pane focus-restore helper reuse, FolderView owned threadpool submit helper reuse, FolderView thumbnail stat helper reuse, FolderView pending-to-paint metric shape reuse, IconCache path failure-store duplicate-race contracts, Microsoft Drive validated Graph/preauthenticated-upload URL types, redirect suppression, redacted diagnostics, authorization-header separation, server upload acknowledgement, safe merge/commit state, shared Microsoft/Google/S3 continuation bounds, Google response/deadline/retry/single-flight/identity guards, transactional plugin configuration/secret scrubbing, shared Curl/Google runtime ownership, cancellable bounded 7z indexing, MTP RAII owner/directory-info handoff contracts, LocalSearch snapshot temp-path exception logging contracts, SearchAndIndex callback exception logging contracts, DxUi native text-input bounded formatting contracts, DxUi focus-sensitive Win32 assertion desktop-probe contracts, FileSystemMtp pragma-warning rationale contracts, shared ordinal string helper usage, shared truthy env-flag helper usage for FolderView WARP/perf flags, shared ViewerSpace opt-in env-flag parser usage, shared DxUi modal loop/quit propagation for archive prompts, FolderView perf-budget strict-mode contracts, FolderView overlay perf advisory sample-sufficiency contract, and Riptide/Floodgate/Observatory source contracts |
| `Tools\Tests\TestInventory.Tests.ps1` | Source-derived test inventory manifest; native project/run-plan parity; source-contract classification and replacement queue; FileOperations phase-order drift guard; and doc-count lint |
| `Tools\Tests\ThemeDistributionContracts.Tests.ps1` | Theme/license distribution, canonical app-theme resolution, semantic-key authoring, named palette profiles, and rejection of exact palette pass-through adapters |
| `Tools\Tests\ViewerChromeSourceContracts.Tests.ps1` | Source/spec guards for shared viewer combo keyboard routing, Unicode clipboard use, first-party title-bar policy, Escape focus-cancel-close docs, and the single detached-console launcher contract |
| `Tools\Tests\VcpkgInstallSafety.Tests.ps1` | vcpkg triplet leaf-name validation and staging/install child path containment |
| `Tools\Tests\Versioning.Tests.ps1` | Local build-number reuse/allocation |
| `Tools\Tests\WingetValidation.Tests.ps1` | Winget validation warning suppression, failure propagation, portable manifest metadata, single detached-console WinGet launcher contracts, and VC runtime ZIP helper coverage |
| `Tests\vcpkg-merge-synthetic-test.ps1` | Fast synthetic vcpkg lock/merge cases |
| `Tests\vcpkg-merge-lock-validation.ps1` | Manual-only vcpkg install/lock validation; intentionally excluded from PR CI and `Run-AllTests.ps1 -Suite Full` because it mutates `.build` |

---

## Coverage Matrix

| Area | Commands | Compare | FileOps | Perf | PoC |
|------|----------|---------|---------|------|-----|
| **UI Dialogs** | ✅ 200+ | — | — | — | — |
| **DxUi Framework** | ✅ | — | — | — | ✅ |
| **Preferences** | ✅ ~100 | — | — | — | — |
| **Shortcuts** | ✅ 30 | — | — | — | — |
| **Navigation** | ✅ 15 | — | — | — | — |
| **Find/Search UI** | ✅ 46 | — | — | — | — |
| **Compare Engine** | — | ✅ 35+ | — | — | — |
| **Search Backends** | — | ✅ 50+ | — | — | — |
| **SQLite Index** | — | ✅ 6 | — | — | — |
| **Search Service** | — | ✅ 30+ | — | — | — |
| **File Copy/Move** | — | — | ✅ 30+ | — | — |
| **File Delete** | — | — | ✅ 10+ | — | — |
| **Conflict Prompts** | — | — | ✅ 7 | — | — |
| **Cross-FS Bridge** | — | — | ✅ 8 | — | — |
| **Remote Storage** | — | ✅ 10 | ✅ 19 | — | — |
| **Bandwidth Throttle** | — | — | ✅ 4 | — | — |
| **Icon Cache** | — | — | — | ✅ | — |
| **Folder View Perf** | — | — | — | ✅ | — |
| **ETW Tracing** | — | — | — | — | ✅ |
| **Viewer Plugins** | — | — | — | — | ✅ |
| **Settings Roundtrip** | ✅ 10 | ✅ 2 | — | — | — |
| **Theme System** | ✅ 20+ | — | — | — | ✅ |
| **OAuth/Credentials** | — | ✅ 5 | — | — | — |
