# FolderView WarpDrive Continuation Baton - 2026-06-29

## Purpose

Archive the exact continuation state for `Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md` after the FolderView perf matrix went green but the final Full suite remained red.

This is a pause/handoff note, not closeout. Do not move the WarpDrive plan to `Specs/Plans/Done/` until the remaining Full-suite blockers are fixed and the final evidence is archived.

## Latest Pause Pointer

Latest continuation archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_151949_folderview_warpdrive_pause\CONTINUATION.md
```

Current blocker at that pause:

- The exact 107-case FileOps+Navigation+Dialogs slice now fails earlier at
  `cmd_pane_navigation_change_directory_keeps_navigation_shell_stable`.
  Evidence: `.build\logs\commands-fileops-navigation-dialogs-to-createdir-rootdiag-20260701_151636.out.log`
  and `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_151702`
  (42 passed / 1 failed / 64 skipped).
- Focused Change Directory passes:
  `.build\logs\commands-change-directory-focused-20260701_151804.out.log` and
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_151806` (1 passed / 0
  failed / 0 skipped), so the next investigation is order-sensitive leakage from
  predecessor navigation/menu/history cases.
- This compact delta archive includes current status, full tracked patch,
  latest logs, the exact 107-case filter, the failing/passing durable Commands
  archives, spec snapshots, manifest, and checksums. The previous heavy archive
  at `Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_151116_folderview_warpdrive_pause\CONTINUATION.md`
  remains the dump bundle for the TSF crash context.
- A Release `RedSalamander.exe` process was running at archive time. Check and
  stop/wait for app processes before running selftests.
- Resume by running the Change Directory predecessor pair/triple first. If the
  minimized cluster does not reproduce, add diagnostic-only fields to
  `TestChangeDirectoryKeepsNavigationShellStable`, rebuild Debug, and rerun the
  exact 107-case filter. Copy any fresh `last_run`, WER events, and dumps before
  launching another `Run-AllTests.ps1`.

## Current State

- Repo: `Z:\src\RedSalamander`
- Branch: `codex/folderview-warpdrive`
- Worktree: intentionally dirty. Do not reset or checkout away these changes.
- Related DxUi/UIA baton: `Specs\Plans\WIP\DxUi_Uia_ContinuationBaton_2026-06-29.md`

Current dirty-state snapshots:

```text
Specs\TestRuns\4cb089111a23\Full\2026-06-29_220021_full_suite_blocked\git-status-short.txt
Specs\TestRuns\4cb089111a23\Full\2026-06-29_220021_full_suite_blocked\git-diff-stat.txt
```

Important modified areas:

- FolderView perf hot paths: `RedSalamander\FolderView.Icons.cpp`, `RedSalamander\FolderView.Rendering.cpp`, `RedSalamander\SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp`
- Tool/test inventory closeout: `Tools\Tests\RunAllTestsPlan.Tests.ps1`, `Tools\Tests\TestInventory.Tests.ps1`, `Specs\Testing\Testing_TestCoverage.md`, `Tests\README.md`
- Perf validation contract: `Specs\Testing\Testing_PerformanceValidation.md`
- Full-suite blocker neighborhood: `Common\DxUi\DxUi.NativeTextInput.cpp`, `Common\DxUi\DxUi.WindowHost.cpp`, `Common\DxUi\DxUi.Accessibility.cpp`, related DxUi/UIA and Commands helper files

## Durable Archive

Latest blocked Full-suite evidence was copied out of volatile `last_run` into:

```text
Specs\TestRuns\4cb089111a23\Full\2026-06-29_220021_full_suite_blocked\
```

Contents:

- `run-all-tests-results.json` - Full suite JSON result.
- `runall-full-skipbuild-20260629_220021_691.out.log` - Full suite stdout log.
- `runall-full-skipbuild-20260629_220021_691.err.log` - Full suite stderr log, empty.
- `ViewerPETests.output.log` - ViewerPE failure context from the Full run.
- `pester-inventory-focused-20260629_224800.log` - focused Pester inventory rerun, now green.
- `crash-dumps-manifest.json` - LocalDumps paths and sizes for the latest RedSalamander crashes.
- `wer-events-redsalamander.json` - Application Error / WER events for the latest crashes.
- `git-status-short.txt`, `git-diff-stat.txt` - worktree snapshot.

Crash dumps were not copied into the repo because each is about 110-145 MB. The manifest records their exact local paths under:

```text
C:\Users\eric\AppData\Local\CrashDumps\
```

## Green Evidence

FolderView perf is green in both Release and Debug on this worktree:

- Test-enabled Release build passed: `.build\logs\msbuild-20260629_210405_389.log`
- Release full FolderView perf matrix passed: `Specs\TestRuns\4cb089111a23\Commands\2026-06-29_210835`
  - `folder.frame.total_us`: count 1709, p50 2862us, p95 7956us, p99 8328us, max 20107us
  - `folder.frame.present_us`: count 1709, p50 340us, p95 7000us, p99 7552us, max 8152us
- Debug build passed: `.build\logs\msbuild-20260629_210918_769.log`
- Debug full FolderView perf matrix passed: `Specs\TestRuns\4cb089111a23\Commands\2026-06-29_211316`
  - `folder.frame.total_us`: count 1655, p50 4122us, p95 23179us, p99 28212us, max 47462us
  - `folder.frame.present_us`: count 1655, p50 348us, p95 21069us, p99 26481us, max 45861us
- Relayout focused evidence passed: `Specs\TestRuns\4cb089111a23\Commands\2026-06-29_210027`
  - `folder.relayout_to_paint_us`: 106600, 72368, 74427, 86725, 95263, 111440

Build/test plumbing fixed during this continuation:

- Test-enabled Debug full build passed: `.build\logs\msbuild-20260629_215806_212.log`
- Monitor ETW latency suite passed inside the later Full run, confirming the `RSBuildEnableTests` propagation fix in `Tools\Tests\RedSalamanderPluginDeployment.Tests.ps1`.
- Focused Pester inventory subset now passes:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Import-Module Pester -ErrorAction Stop; Invoke-Pester -Path @('Z:\src\RedSalamander\Tools\Tests\RunAllTestsPlan.Tests.ps1','Z:\src\RedSalamander\Tools\Tests\TestInventory.Tests.ps1') -EnableExit"
```

Evidence:

```text
Specs\TestRuns\4cb089111a23\Full\2026-06-29_220021_full_suite_blocked\pester-inventory-focused-20260629_224800.log
Passed: 13 Failed: 0 Skipped: 0 Pending: 0 Inconclusive: 0
```

## Latest Blocked Full Suite

Command:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild
```

Evidence:

```text
Specs\TestRuns\4cb089111a23\Full\2026-06-29_220021_full_suite_blocked\run-all-tests-results.json
Specs\TestRuns\4cb089111a23\Full\2026-06-29_220021_full_suite_blocked\runall-full-skipbuild-20260629_220021_691.out.log
```

Result:

- Exit code: 1
- Total: 206
- Passed: 175
- Failed: 2
- Skipped: 29
- Duration: 1,194,565 ms

Failing/blocking surfaces:

- `Commands`: SelfTest crashed, exit `-1073741819` (`0xC0000005`), duration 753,138 ms, no result cases emitted.
- `FileOperations`: SelfTest crashed, exit `-1073741819` (`0xC0000005`), duration 7,070 ms, no result cases emitted.
- `ViewerPETests`: executable exited 1. Full log reported a ViewerVLC HUD focus failure earlier; archived `ViewerPETests.output.log` ends with a fresh-harness pass followed by `ViewerPETests failed`, so rerun this suite focused before changing code.
- `ToolsPesterTests`: Full run exited 1, but the focused inventory subset that failed in the Full log has since been fixed and rerun green. Full Tools Pester still needs rerun after the native crash blockers are handled.

## 2026-06-30 Resume Update

Fresh broad Commands rerun no longer reproduced the earlier Connection Manager/TSF crash. It progressed through Connection Manager and Preferences, then failed a specific retained-search focus case:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2
```

## 2026-06-30 12:05 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-06-30_120509_folderview_warpdrive_pause\
```

Resume from this archive, not the 11:26 archive. It includes the current tracked
patch, git status/name-status/stat, untracked list, plan/baton snapshots, copied
latest `.build\codex-runs\*_20260630_resume3*` evidence, the volatile
`last_run` directory, a manifest, and SHA256 checksums.

The 11:26 Connection Manager tab-traversal blocker no longer reproduces in the
focused, paired, block, plugin-config-plus-connection-manager, or prefix runs.
The active broad-suite blocker is now:

```text
cmd_compare_directories_options_live_dx_body_interaction
Reason: Compare Directories options DX edit 'Ignore files' did not restore its original value.
Broad evidence: .build\codex-runs\commands_broad_after_connection_prefix_green_20260630_resume3
Broad result: 519 passed / 1 failed / 256 skipped
```

Focused Compare Options and the immediate Compare Options cluster pass. The best
current repro is:

```text
.build\codex-runs\commands_prefs_compare_then_compare_options_20260630_resume3
Result: 6 passed / 1 failed / 0 skipped
```

Resume by adding diagnostics around the failing `Ignore files` edit restore path
in `RedSalamander\SelfTest\Commands\Commands.SelfTest.CompareOptions.cpp`.
Capture initial, edited, final observed, and all matching visible edit states
before changing production code. Then prove whether the root cause is
Preferences Compare Directories deferred focus restoration, retained settings
state, or wrong/stale UIA edit matching.

Do not run multiple `Run-AllTests.ps1` invocations in parallel; they share
`C:\Users\eric\AppData\Local\RedSalamander\SelfTest\last_run` and corrupt each
other's archive evidence.

## 2026-06-30 12:36 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-06-30_123644_folderview_warpdrive_pause\
```

Resume from this archive, not the 12:05 archive. It includes the current tracked
patch, git status/name-status/stat, untracked list, plan/baton snapshots, copied
latest `.build\codex-runs` evidence, selected volatile `last_run` files, a
continuation note, and SHA256 checksums.

New evidence since 12:05:

- Debug build passed with the Compare Options diagnostics:
  `Z:\src\RedSalamander\.build\logs\msbuild-20260630_121407_554.log`
  (0 warnings / 0 errors).
- The previous Compare Options `Ignore files` blocker did not reproduce:
  `.build\codex-runs\commands_prefs_compare_then_compare_options_diag_20260630`
  (7 passed / 0 failed / 0 skipped).
- Broad Commands now fails later:
  `.build\codex-runs\commands_broad_after_compare_diag_20260630`
  (565 passed / 1 failed / 210 skipped).
- The active broad-suite blocker is
  `cmd_pane_selection_select_unselect_dialogs_keep_navigation_shell_stable`:
  `Navigation shell did not stay quiet during Unselect dialog cancel; focusTarget=0,
  editMode=no, historyVisible=no, suggestVisible=no, popupVisible=no,
  childWindows=0, currentPath='', historyCount=0, refreshCount=21, itemCount=3,
  selectedCount=3, focusedItem='b.log'.`
- Focused selection/unselect and the immediate 13-case selection cluster are green:
  `.build\codex-runs\commands_selection_unselect_navshell_focused_20260630`
  and `.build\codex-runs\commands_selection_navshell_cluster_553_565_20260630`.

Current root-cause lead:

- Treat the selection/unselect failure as order-sensitive until proven otherwise.
  Focused and immediate-cluster runs pass.
- The failure message lacks baseline refresh count, expected root, baseline
  history count, snapshot-validity state, actual focus HWND, and focused pane.
  Add diagnostics in
  `RedSalamander\SelfTest\Commands\Commands.SelfTest.Navigation.cpp` before
  changing product behavior.
- Inspect `WaitForNavigationViewSnapshot` and `DebugGetNavigationViewSnapshot`
  to determine whether `currentPath=''` means a failed snapshot or a real blank
  navigation path.

Recommended next commands:

```powershell
rg -n "WaitForNavigationViewSnapshot\(" RedSalamander\SelfTest\Commands\Commands.SelfTest.Settings.cpp RedSalamander\SelfTest\Commands\Commands.SelfTest.Navigation.cpp RedSalamander\SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp
rg -n "DebugGetNavigationViewSnapshot|DebugGetForceRefreshCount|TestPaneSelectionSelectUnselectDialogsKeepNavigationShellStable" RedSalamander Common
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_pane_selection_select_unselect_dialogs_keep_navigation_shell_stable
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter navigation_location_edit_input_expands_environment_variables,cmd_pane_selection_goto_selected_name,cmd_pane_selection_goto_selected_name_keeps_navigation_shell_stable,cmd_pane_selection_select_same_extension_keeps_navigation_shell_stable,cmd_pane_selection_select_same_name_keeps_navigation_shell_stable,cmd_pane_selection_select_next_keeps_navigation_shell_stable,cmd_pane_selection_select_calculate_directory_size_next_keeps_navigation_shell_stable,cmd_pane_selection_select_all_keeps_navigation_shell_stable,cmd_pane_selection_invert_keeps_navigation_shell_stable,cmd_pane_selection_save_restore_keeps_navigation_shell_stable,cmd_pane_selection_hide_show_names_keeps_navigation_shell_stable,cmd_pane_selection_hide_unselected_names_keeps_navigation_shell_stable,cmd_pane_selection_select_unselect_dialogs_keep_navigation_shell_stable
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2
```

Evidence:

```text
Specs\TestRuns\4cb089111a23\Commands\2026-06-30_081705\
```

Result:

- Exit code: 1
- Duration: 279,895 ms
- Passed: 284
- Failed: 1
- Skipped: 491
- Failing case: `cmd_preferences_dialog_themes_search_roundtrip_preserves_retained_state`
- Failure reason: `Failed to refocus the Preferences category host before leaving Themes for General during retained-search round-trip validation.`

Useful files in the archive:

- `commands_results.json` - structured case result payload.
- `commands_trace.txt` - concise Commands trace; failure is at lines 689-694.
- `selftest_run_results.json` / `selftest_run_trace.txt` - combined selftest payloads.
- `runall-commands-failfast-20260630_081224.out.log` - captured runner stdout from `.build\codex-runs\commands_broad_20260630_081224`.
- `runall-commands-failfast-20260630_081224.err.log` - empty stderr.
- `git-status-short.txt`, `git-diff-stat.txt` - worktree snapshot at archive time.

The exact Connection Manager case that appeared in an older crash stack was rerun focused and passed:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_connection_manager_window_uses_dxui_command_buttons
```

Evidence:

```text
Specs\TestRuns\4cb089111a23\Commands\2026-06-30_080925
```

No new RedSalamander crash report or LocalDump was created during the 2026-06-30 Commands rerun. The old `textinputframework.dll` dump evidence below remains relevant for FileOperations/native teardown investigation, but the immediate Commands blocker is now the Preferences category-host focus failure.

## Crash Evidence

Focused `Commands` rerun reproduced the native crash after the Full run:

- Dump: `C:\Users\eric\AppData\Local\CrashDumps\RedSalamander.exe.97872.dmp`
- WER time: 2026-06-29 22:35:30 Europe/Paris
- Faulting module: `textinputframework.dll`
- Exception: `0xC0000005`
- Fault offset: `0x000000000006a419`
- Report Id: `6ce2cd72-d03b-4eef-90c7-4fa7899dd888`

Other same-session crash dumps in the manifest include:

- `C:\Users\eric\AppData\Local\CrashDumps\RedSalamander.exe.48056.dmp` - textinputframework crash at 22:14:46.
- `C:\Users\eric\AppData\Local\CrashDumps\RedSalamander.exe.106368.dmp` - ntdll crash at 22:14:56.
- `C:\Users\eric\AppData\Local\CrashDumps\RedSalamander.exe.74052.dmp` - textinputframework crash at 21:49:35.
- `C:\Users\eric\AppData\Local\CrashDumps\RedSalamander.exe.86104.dmp` - ntdll crash at 21:49:45.

No command-line debugger was available on PATH during this session (`cdb`, `windbg`, `dumpchk` not found). Visual Studio 18 is installed under `C:\Program Files\Microsoft Visual Studio\18`; use Visual Studio's dump debugger or install Windows Debugging Tools to get a stack before applying a fix.

## Current Crash Hypothesis To Verify

Do not treat this as proven. The strongest lead is `Common\DxUi\DxUi.NativeTextInput.cpp` TSF lifetime churn:

- Current dirty code changed `DeactivateNativeTextInputTsf()` to call `ITfThreadMgr::SetFocus(nullptr)` before `ITfDocumentMgr::Pop(TF_POPF_ALL)`.
- Current dirty code disconnects the text store after focus/pop instead of before document pop.
- Current dirty code calls `ShutdownNativeTextInputThreadManagerState(GetNativeTextInputThreadManagerState())` on every active TSF deactivation.
- HEAD order was: disconnect text store, `documentMgr->Pop(TF_POPF_ALL)`, `threadMgr->SetFocus(nullptr)`, release per-host COM pointers, and keep the thread manager alive until process/thread shutdown.

The repeated `textinputframework.dll` offset `0x6a419` and paired `ntdll.dll` crash after broad suite churn make this the first place to inspect, but get a dump stack if possible.

Minimal test plan for that hypothesis:

1. Open `RedSalamander.exe.97872.dmp` in Visual Studio or cdb and confirm the crashing thread stack.
2. If the stack is in TSF/text input teardown, make the smallest patch in `DeactivateNativeTextInputTsf()` only:
   - restore disconnect-before-pop,
   - restore pop-before-SetFocus-null,
   - remove per-deactivation `ShutdownNativeTextInputThreadManagerState(...)`,
   - keep `ShutdownNativeTextInputForCurrentThread()` for real process/thread teardown.
3. Run:

```powershell
.\build.ps1 -ProjectName DxUiTests -Configuration Debug
.\.build\x64\Debug\DxUiTests.exe
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2
.\Tools\Run-AllTests.ps1 -Suite FileOperations -SkipBuild -FailFast -TimeoutMultiplier 2
```

4. If those pass, rerun the blocked surfaces:

```powershell
.\.build\x64\Debug\ViewerPETests.exe
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Import-Module Pester -ErrorAction Stop; Invoke-Pester -Path 'Z:\src\RedSalamander\Tools\Tests' -EnableExit"
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild
```

## Pester Inventory Work Already Applied

These changes are already in the dirty worktree and focused-green:

- `Tools\Tests\RunAllTestsPlan.Tests.ps1`
  - Full suite expected plan now includes `SettingsSchemaTests` and `CrashHandlingTests`.
- `Tools\Tests\TestInventory.Tests.ps1`
  - Commands static registrations: 674.
  - CompareDirectories static registrations: 184.
  - FileOperations active phases: 120.
  - NativeTextInput cases: 117.
  - Tools Pester cases: 108.
- `Specs\Testing\Testing_TestCoverage.md`
  - Source-derived counts updated.
  - Runner-native header kept explicitly as the last 2026-06-23 snapshot so it does not claim today's source-derived active phase count.
- `Tests\README.md`
  - NativeTextInput and Tools Pester counts updated.

## FolderView Work To Preserve

These are the FolderView pieces that made the perf matrix pass and must not be lost:

- `RedSalamander\FolderView.Icons.cpp`
  - Removed file-backed `SelfTest::AppendSelfTestTrace(...)` from the icon update hot path.
- `RedSalamander\FolderView.Rendering.cpp`
  - Removed file-backed `SelfTest::AppendSelfTestTrace(...)` from name-filter `Present1` render path.
- `RedSalamander\SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp`
  - Added quiescence/warm-snapshot helpers.
  - Hardened quick-search input helpers.
  - Added relayout phase metrics for DPI/theme/SetWindowPos/pump/redraw/warm render.
  - Stabilized the full FolderView perf matrix in Debug and Release.
- `Specs\Testing\Testing_PerformanceValidation.md`
  - Added the durable rule banning file-backed `SelfTest::AppendSelfTestTrace(...)` in measured render/draw/icon/present hot paths.

## Remaining Closeout

1. Fix the current broad Commands blocker:
   `cmd_preferences_dialog_panes_history_size_live_dx_interaction`.
2. Rerun the broad Commands gate.
3. Rerun FileOperations focused/broad; if it still crashes, use the old TSF/LocalDump evidence plus a real dump stack before patching native text input teardown.
4. Rerun `ViewerPETests.exe` focused and fix or quarantine the ViewerVLC HUD focus failure with evidence.
5. Rerun full `Tools\Tests` Pester, not only the inventory subset.
6. Rerun `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild`.
7. If Full is green, rerun final FolderView Debug and test-enabled Release perf matrices or cite fresh enough green archives from this baton if no code changed in FolderView.
8. Update `Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md` with final evidence.
9. Move the WarpDrive plan to `Specs\Plans\Done\` only after all gates are green.

## 2026-06-30 Pause Archive Refresh

Dedicated pause archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-06-30_083849_folderview_warpdrive_pause\
```

That folder preserves:

- current `git status --short --branch`,
- current `git diff --stat`,
- full tracked dirty patch as `git-diff.patch`,
- untracked-file list,
- pre-refresh copies of this baton and the Operation plan,
- latest broad Commands runner stdout/stderr,
- latest broad Commands result JSON and trace.

Latest Commands state:

- The earlier Themes retained-search failure did not reproduce:
  - focused Themes: `Specs\TestRuns\4cb089111a23\Commands\2026-06-30_082128` (1 passed / 0 failed),
  - Keyboard+Themes pair: `Specs\TestRuns\4cb089111a23\Commands\2026-06-30_082234` (2 passed / 0 failed),
  - Preferences prefix: `Specs\TestRuns\4cb089111a23\Commands\2026-06-30_082834` (173 passed / 0 failed).
- Broad Commands rerun failed later at:

```text
Specs\TestRuns\4cb089111a23\Commands\2026-06-30_083421
```

Result:

- Passed: 311
- Failed: 1
- Skipped: 464
- Failing case: `cmd_preferences_dialog_panes_history_size_live_dx_interaction`
- Failure reason: `Preferences Panes page did not settle to the active DX surface before history-size validation.`

Focused Panes repro attempts passed:

- `cmd_preferences_dialog_panes_history_size_live_dx_interaction`: `Specs\TestRuns\4cb089111a23\Commands\2026-06-30_083612` (1 passed / 0 failed).
- `cmd_preferences_dialog_panes_tab_traversal_live_dx_interaction,cmd_preferences_dialog_panes_history_size_live_dx_interaction`: `Specs\TestRuns\4cb089111a23\Commands\2026-06-30_083633` (2 passed / 0 failed).

Current root-cause lead:

- `RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.ThemesGeneralPanes.cpp`
  contains a brittle `navigateToPanesPage` helper in the Panes history-size test.
- It focuses the category tree and sends one `VK_DOWN`, which assumes the retained Preferences category is immediately before Panes.
- The failure only appears in broad order, after many previous Preferences cases have changed retained category/page state.
- The adjacent Panes tab-traversal helper uses the same relative navigation assumption and should be considered for the same hardening while the file is open.

Recommended first patch after resume:

1. Replace the relative one-step `VK_DOWN` Panes navigation with deterministic `DebugSelectPreferencesCategory(kPrefCategoryPanes)`.
2. Keep the existing settle predicate that requires `currentCategory == kPrefCategoryPanes` and an active DX surface.
3. Build `RedSalamander` Debug.
4. Rerun:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_preferences_dialog_panes_history_size_live_dx_interaction
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter 'cmd_preferences_dialog_panes_tab_traversal_live_dx_interaction,cmd_preferences_dialog_panes_history_size_live_dx_interaction'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2
```

No `RedSalamander.exe` process was left running at the pause archive point.

## 2026-06-30 Continuation Archive Refresh

Dedicated continuation archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-06-30_092000_folderview_warpdrive_continuation\
```

That folder preserves:

- `git-status-short.txt`, `git-diff-stat.txt`, `git-diff-name-status.txt`, and `untracked-files.txt`;
- full tracked dirty patch as `patches\tracked-working-tree.diff`;
- current copies of this baton, the DxUi/UIA baton, and the WarpDrive WIP plan under `specs\`;
- latest broad Commands stdout/stderr and `last_run\commands\trace.txt` under `logs\`;
- WER text, `Report.wer`, crash-dump manifest, and a copied local dump under `crash\RedSalamander.exe.43200.dmp`.

Panes broad-order blocker resolved:

- Patched `RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.ThemesGeneralPanes.cpp` so Panes setup helpers use deterministic `DebugSelectPreferencesCategory(kPrefCategoryPanes)` instead of retained-category-relative one-step keyboard navigation.
- Debug build passed: `.build\logs\msbuild-20260630_084155_463.log`.
- Focused Panes history-size passed: `Specs\TestRuns\4cb089111a23\Commands\2026-06-30_084408`.
- Panes tab-traversal + history-size pair passed: `Specs\TestRuns\4cb089111a23\Commands\2026-06-30_084424`.
- Rebuilt-binary Panes pair passed again after the TSF patch: `Specs\TestRuns\4cb089111a23\Commands\2026-06-30_085941`.

Credential prompt broad-order failure resolved as TSF teardown fallout:

- Broad Commands after the Panes patch failed later at `cmd_connection_credential_prompt_theme_cycle_keeps_surface_legible`: `Specs\TestRuns\4cb089111a23\Commands\2026-06-30_084628`.
- Focused credential case passed: `Specs\TestRuns\4cb089111a23\Commands\2026-06-30_084735`.
- Immediate predecessor + credential pair passed: `Specs\TestRuns\4cb089111a23\Commands\2026-06-30_084750`.
- `cmd_connection_` prefix passed: `Specs\TestRuns\4cb089111a23\Commands\2026-06-30_084857`.
- Patched `Common\DxUi\DxUi.NativeTextInput.cpp` back to the safer TSF teardown order: disconnect text store before TSF callbacks, pop document before clearing TSF focus, and do not shut down the thread-manager state on every deactivation.
- Updated `Tests\DxUiTests\DxUiTests.NativeTextInput.cpp` to guard that teardown order.
- `.\build.ps1 -ProjectName DxUiTests -Configuration Debug` passed: `.build\logs\msbuild-20260630_085136_095.log`.
- `.\.build\x64\Debug\DxUiTests.exe` passed.
- `.\build.ps1 -ProjectName RedSalamander -Configuration Debug` passed: `.build\logs\msbuild-20260630_085545_300.log`.
- Focused credential case passed after patch: `Specs\TestRuns\4cb089111a23\Commands\2026-06-30_085809`.
- `cmd_connection_` prefix passed after patch: `Specs\TestRuns\4cb089111a23\Commands\2026-06-30_085921`.

Current blocker at archive time:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2
```

Run directory:

```text
.build\codex-runs\commands_full_after_tsf_20260630_085959\
```

Result:

- Exit code: `-1073741819` (`0xC0000005`).
- Runtime: about 12m52s.
- No `results.json` was written because the process crashed.
- Trace reached `cmd_pane_pack_prompt_uses_dxui_unique_archive_and_packer_extensions`.
- The run progressed past the earlier Connection Manager, credential prompt, Preferences, Compare Directories, and FileOps issues-pane blockers before the crash.

Crash evidence:

- Dump copied into the continuation archive: `Specs\TestRuns\4cb089111a23\Continuation\2026-06-30_092000_folderview_warpdrive_continuation\crash\RedSalamander.exe.43200.dmp`.
- Original dump path: `C:\Users\eric\AppData\Local\CrashDumps\RedSalamander.exe.43200.dmp`.
- WER report copied into the continuation archive: `crash\Report.wer`.
- WER report id: `f0763cd7-7767-43bc-9e3c-f6b09dc1a7e7`.
- Faulting module: `textinputframework.dll`.
- Exception: `0xC0000005`.
- Fault offset: `0x000000000006a419`.

Next resume checklist:

1. Inspect the copied dump or WER in the continuation archive; confirm whether the stack is still TSF teardown or a different text-input lifecycle path.
2. Search and read the failing case: `cmd_pane_pack_prompt_uses_dxui_unique_archive_and_packer_extensions`.
3. Reproduce focused:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_pane_pack_prompt_uses_dxui_unique_archive_and_packer_extensions
```

4. If focused passes, run predecessor/prefix slices around pane pack prompt to expose broad-order state leakage.
5. Patch only after root cause is proven, then rerun focused, prefix, and broad Commands.
6. Continue with FileOperations, ViewerPETests, Tools Pester, and final Full suite only after broad Commands is green.

## 2026-06-30 09:46 Continuation Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-06-30_094629_folderview_warpdrive_continuation\
```

This archive supersedes the 09:20 archive for resume purposes. It keeps the earlier evidence and adds the focused/minimized crash investigation around the archive pack prompt:

- full tracked dirty patch: `worktree\git_diff_binary.patch`;
- current WIP plan and continuation batons under `plans\`;
- focused build/test logs and copied command evidence under `logs\` and `test-runs\`;
- copied latest crash dump and WER folders under `crash\`;
- `README.md`, `manifest.json`, and `sha256.txt` for quick resume and integrity checks.

Additional minimization evidence now archived:

- focused `cmd_pane_pack_prompt_uses_dxui_unique_archive_and_packer_extensions` passed: `test-runs\2026-06-30_092139`;
- immediate predecessor + pack prompt pair passed: `logs\commands_pack_pair_20260630_092230`;
- `600..643` tail passed: `logs\commands_tail_600_643_20260630_092740`;
- `520..643` tail passed: `logs\commands_tail_520_643_20260630_093000`;
- `400..643` reproduced the native crash: `logs\commands_tail_400_643_20260630_093150`;
- `400..429 + 460..643` reproduced the native crash: `logs\commands_ranges_400_429_460_643_20260630_093920`;
- `400..414 + 460..643` failed earlier at clipboard refocus: `logs\commands_ranges_400_414_460_643_20260630_094420`;
- `415..429 + 460..643` failed earlier at async shortcut clipboard seeding: `logs\commands_ranges_415_429_460_643_20260630_094330`.

Current narrowed blocker:

- broad Commands still crashes in `textinputframework.dll` at offset `0x000000000006a419`;
- focused pack prompt and late tails pass;
- the crash requires the early Find-dialog block `400..429` plus the later modal tail `460..643`;
- latest copied dump: `crash\RedSalamander.exe.78680.dmp`;
- WER report id observed during investigation: `80139524-4da7-4301-a863-ae7bba7ca66a`.

Next resume target:

1. Inspect `Common\DxUi\DxUi.WindowHost.cpp` around `WindowHost::HandleMessage`, especially `WM_NCDESTROY`, focus, and text-input teardown.
2. Inspect `ArchivePackPromptWindow::WndProc` in `RedSalamander\FolderWindow.FileSystem.Commands.Part.cpp` and compare destroy ordering with other modal prompt windows.
3. Inspect DxUi accessibility target/provider teardown in `Common\DxUi\DxUi.Accessibility.cpp`.
4. Prove root cause before behavior changes; then rerun `DxUiTests.exe`, focused pack prompt, `400..429 + 460..643`, broad Commands, and finally the remaining plan gates.

## 2026-06-30 09:59 Continuation Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-06-30_095904_folderview_warpdrive_continuation\
```

This archive supersedes the 09:46 archive for resume purposes while preserving the
09:46 README, manifest, and checksum under `predecessor\`.

New archived evidence:

- full tracked dirty patch before archive creation:
  `worktree\tracked-working-tree-before-archive.patch`;
- current dirty status before and after archive creation under `worktree\`;
- current WIP plan and both continuation batons under `plans\`;
- TSF source-order RED evidence:
  `logs\codex-runs\dxui_native_tsf_red_20260630_095107\`;
- Debug build logs:
  `.build\logs\msbuild-20260630_095108_044.log`,
  `.build\logs\msbuild-20260630_095206_346.log`,
  `.build\logs\msbuild-20260630_095405_958.log`;
- focused pack prompt after TSF focus-clear patch:
  `test-runs\Commands\2026-06-30_095616` and
  `logs\codex-runs\commands_pack_focused_after_tsf_focusclear_20260630_095614\`;
- minimized former crash slice after TSF focus-clear patch:
  `test-runs\Commands\2026-06-30_095708` and
  `logs\codex-runs\commands_ranges_400_429_460_643_after_tsf_focusclear_20260630_095636\`.

Current state at archive time:

- The TSF patch under test changes `WindowHost::DeactivateNativeTextInputTsf()` to
  disconnect the text store, clear TSF focus with `ITfThreadMgr::SetFocus(nullptr)`,
  then pop the document.
- Focused pack prompt passed after the patch.
- The former native crash slice no longer crashes, but fails normally at
  `cmd_pane_find_dialog_command_enablement_matches_idle_running_and_selection_states`
  with reason: `Find split action menu did not highlight the item under a stationary pointer.`
- A fresh archive-time `.\.build\x64\Debug\DxUiTests.exe` run failed at
  `TestStationaryMouseDoesNotOverrideKeyboardRootSwitch` with reason:
  `a stationary mouse over B does not pull keyboard navigation back from A`.
  Earlier in-thread evidence had `DxUiTests.exe` passing immediately after the TSF patch,
  so the next session should re-run it before deciding whether this is environmental,
  order-sensitive, or a real regression.

Next resume target:

1. Re-run `.\.build\x64\Debug\DxUiTests.exe`; if the stationary-pointer menu failure repeats, debug it first.
2. Run focused `cmd_pane_find_dialog_command_enablement_matches_idle_running_and_selection_states`.
3. If focused Find passes, continue order-sensitive minimization of `400..429 + 460..643`.
4. If focused Find fails, debug the split-button/menu stationary-hover path.
5. Re-run the minimized slice from
   `logs\codex-runs\commands_ranges_400_429_460_643_after_tsf_focusclear_20260630_095636\casefilter.txt`.
6. After Commands is green, resume Task 10 Step 4 and Task 11 in the WarpDrive WIP plan:
   Debug + test-enabled Release FolderView perf matrix, Full suite, authoritative spec closeout, then move to Done.

## 2026-06-30 10:17 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-06-30_101733_folderview_warpdrive_pause\
```

Resume from this archive, not the earlier 09:59 one. The archive includes:

- `CONTINUATION.md` with the exact next debug checklist;
- `git-status-short.txt`, `git-diff-stat.txt`, `focused-diff.patch`, `full-dirty-diff.patch`, and `untracked-files.txt`;
- snapshots of the WIP plan and continuation batons;
- copied focused and minimized Commands artifacts;
- copied `.build\codex-runs` logs for DxUi, focused Find, the minimized slice, and the current focused blocker.

New evidence since the 09:59 archive:

- `.\.build\x64\Debug\DxUiTests.exe` passed. Evidence copied from `.build\codex-runs\dxui_stationary_pointer_repro_20260630_resume\`.
- Focused `cmd_pane_find_dialog_command_enablement_matches_idle_running_and_selection_states` passed after selftest cursor-settle hardening. Evidence: `Specs\TestRuns\4cb089111a23\Commands\2026-06-30_101418`.
- The minimized former crash slice `400..429 + 460..643` now gets past Find and fails at `cmd_pane_clipboardPaste_ignores_unfocused_navigation_edit`. Evidence: `Specs\TestRuns\4cb089111a23\Commands\2026-06-30_101514`.
- Focused `cmd_pane_clipboardPaste_ignores_unfocused_navigation_edit` also fails. Evidence: `Specs\TestRuns\4cb089111a23\Commands\2026-06-30_101544`.

Current blocker:

```text
cmd_pane_clipboardPaste_ignores_unfocused_navigation_edit
Reason: Failed to refocus left folder view after opening the navigation edit field.
```

Root-cause note:

- `WaitForFolderViewPaneFocus(...)` already calls `FocusFolderViewPane(pane)` before and during the wait.
- The failure is therefore not merely a missing test call to focus the pane.
- Resume by inspecting the interaction between `TestPaneClipboardPasteIgnoresUnfocusedNavigationEdit(...)`, `FolderWindow::CommandChangeDirectory(...)`, the navigation edit session, and pane focus bookkeeping (`GetFocus()`, `GetFocusedFolderViewHwnd()`, `GetFocusedPane()`).

Recommended next commands:

```powershell
rg -n "TestPaneClipboardPasteIgnoresUnfocusedNavigationEdit|WaitForFolderViewPaneFocus|FocusFolderViewPane|CommandChangeDirectory|TryHandleNavigationEditClipboardCommand" RedSalamander Common
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_pane_clipboardPaste_ignores_unfocused_navigation_edit
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter "400..429,460..643"
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2
```

## 2026-06-30 11:26 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-06-30_112600_folderview_warpdrive_pause\
```

Resume from this archive, not the 11:09 archive. It includes the current tracked
patch, git status/name-status/stat, untracked list, plan/baton snapshots, copied
`.build\codex-runs` evidence, trace excerpts, a manifest, and SHA256 checksums.

New evidence since the 11:09 archive:

- The exact last160-plus-target prefix is now green:
  `.build\codex-runs\commands_filter_prompt_prefix_last160_freshred_20260630_resume2`
  (161 passed / 0 failed / 0 skipped).
- Broad Commands then failed at
  `cmd_connection_credential_prompt_long_run_open_close_stays_stable`, but the
  focused long-run and immediate credential-prompt cluster are green:
  `.build\codex-runs\commands_credential_prompt_longrun_focused_20260630_resume2`
  and `.build\codex-runs\commands_credential_prompt_cluster_20260630_resume2`.
- The current broad-suite blocker is
  `cmd_connection_manager_window_tab_traversal_live_dx_interaction`:
  `.build\codex-runs\commands_broad_credential_recheck_20260630_resume2`
  (203 passed / 1 failed / 572 skipped).

Current blocker:

```text
cmd_connection_manager_window_tab_traversal_live_dx_interaction
Reason: Connection Manager did not move focus to the Protocol combo during keyboard traversal validation.
Trace lead: preKind='Edit' preLabel='Name' preModifiers=0x4; routed Tab lands on CommandButton 'Remove'; postModifiers=0x0.
```

Root-cause note:

- Treat `preModifiers=0x4` as a lead only. Prove whether this is modifier-state
  residue, reverse traversal, native text focus behavior, or tab-order/focus
  bookkeeping before changing production or test code.
- Do not start by patching clipboard/navigation-edit or credential-prompt code;
  those failures are not the current reproducible red point.

Recommended next commands:

```powershell
rg -n "TestConnectionManagerWindowTabTraversalLiveDxInteraction|ConnectionManager tab traversal|preModifiers|postModifiers|VK_SHIFT|modifier" RedSalamander\SelfTest\Commands Common\DxUi RedSalamander\ConnectionManagerWindow.cpp
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_connection_manager_window_tab_traversal_live_dx_interaction
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_connection_manager_window_long_run_open_close_stays_stable,cmd_connection_manager_window_tab_traversal_live_dx_interaction
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2
```

## 2026-06-30 13:01 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-06-30_130129_folderview_warpdrive_pause\
```

Resume from this archive, not the 12:36 archive. It includes the current tracked
patch, git status/name-status/stat, untracked list, plan/baton snapshots, the
latest Debug build log, copied `.build\codex-runs` evidence, a manifest,
`CONTINUATION.md`, and SHA256 checksums.

New evidence since the 12:36 archive:

- Debug build passed after selection/unselect diagnostic context was added:
  `Z:\src\RedSalamander\.build\logs\msbuild-20260630_124148_516.log`
  (0 warnings / 0 errors).
- The previous selection/unselect blocker did not reproduce:
  `.build\codex-runs\commands_selection_unselect_navshell_focused_diag_20260630`
  (1 passed / 0 failed / 0 skipped),
  `.build\codex-runs\commands_selection_navshell_cluster_553_565_diag_20260630`
  (13 passed / 0 failed / 0 skipped),
  `.build\codex-runs\commands_fileops_issues_to_selection_diag_20260630`
  (35 passed / 0 failed / 0 skipped), and
  `.build\codex-runs\commands_shortcuts_compare_fileops_selection_diag_20260630`
  (66 passed / 0 failed / 0 skipped).
- Broad Commands now fails earlier:
  `.build\codex-runs\commands_broad_after_selection_diag_20260630`
  (381 passed / 1 failed / 394 skipped).

Current blocker:

```text
cmd_preferences_dialog_category_tree_handles_reverse_keyboard_navigation
Reason: Preferences category host lost keyboard focus after the second reverse-navigation VK_UP.
```

Focused and predecessor repro attempts are green:

```text
.build\codex-runs\commands_prefs_category_reverse_focused_20260630
.build\codex-runs\commands_prefs_compare_to_reverse_cluster_20260630
.build\codex-runs\commands_prefs_350_to_reverse_cluster_20260630
```

Root-cause note:

- Treat this as broad-order residue until proven otherwise.
- The failing test currently uses raw `SetFocus(categoryTreeHost)`, sends
  `VK_END`, then two `VK_UP` messages, and immediately asserts
  `snapshot.categoryTreeFocused`.
- The first resume patch should add diagnostic context to
  `TestPreferencesDialogCategoryTreeReverseKeyboardNavigation` in
  `RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp`
  before changing product behavior.

Recommended next commands:

```powershell
rg -n "TestPreferencesDialogCategoryTreeReverseKeyboardNavigation|categoryTreeFocused|DebugGetPreferencesDialogSnapshot|FocusWindowAndWait" RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp RedSalamander\Preferences.h RedSalamander\Preferences.Dialog.cpp RedSalamander\Preferences.cpp RedSalamander\SelfTest\Commands\Commands.SelfTest.Settings.cpp
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_preferences_dialog_category_tree_handles_reverse_keyboard_navigation
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_preferences_dialog_compare_directories_page_uses_dxui_statics,cmd_preferences_dialog_compare_directories_live_dx_interaction,cmd_preferences_dialog_compare_directories_content_workers_ignore_files_live_dx_interaction,cmd_preferences_dialog_compare_directories_tail_toggles_live_dx_interaction,cmd_preferences_dialog_compare_directories_tab_traversal_live_dx_interaction,cmd_preferences_dialog_compare_directories_roundtrip_restores_dxui_surface,cmd_preferences_dialog_compare_directories_theme_cycle_keeps_surface_legible,cmd_preferences_dialog_category_tree_keyboard_navigation_updates_category,cmd_preferences_dialog_editors_mouse_live_dx_notes,cmd_preferences_dialog_editors_mouse_tab_skips_note_surface,cmd_preferences_dialog_viewers_editors_file_action_settings_apply,cmd_preferences_dialog_category_tree_handles_reverse_keyboard_navigation
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2
```

After broad Commands is green, continue with FileOperations, ViewerPETests,
Tools Pester, Full, final FolderView perf evidence, and plan move-to-Done only
after every gate is green.

## 2026-06-30 22:15 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-06-30_221502_folderview_warpdrive_pause\CONTINUATION.md
```

Resume from this archive, not the older blocker notes above. It contains the
current tracked patch, git status/stat/numstat, untracked list, copied test
artifacts, copied build logs, crash-dump manifest, process snapshot, and WIP
spec snapshots.

New evidence since the older baton state:

- Broad Commands is green after the TSF teardown-order fix:
  `.build\codex-runs\commands_broad_after_tsf_teardown_order_20260630`
  (774 passed / 0 failed / 2 skipped).
- Debug build after TSF teardown-order fix passed:
  `Z:\src\RedSalamander\.build\logs\msbuild-20260630_211155_814.log`.
- Focused FileOps crash repro is green after replacing synchronous popup
  `RedrawWindow(... RDW_UPDATENOW)` with asynchronous `InvalidateRect(...)`:
  `.build\codex-runs\fileops_phase1_async_popup_invalidate_20260630`
  (3 passed / 0 failed / 0 skipped).
- Focused Floodgate selftest hook is green after normalizing `/` and `\`
  separators inside the `ENABLE_TESTS` bridge mutation hook:
  `.build\codex-runs\fileops_floodgate_move_cleanup_corruption_normalized_hook_20260630`
  (3 passed / 0 failed / 0 skipped).
- Full FileOps is green:
  `.build\codex-runs\fileops_full_after_floodgate_hook_normalization_20260630`
  (102 passed / 0 failed / 20 skipped).
- Debug builds after the FileOps fixes passed:
  `Z:\src\RedSalamander\.build\logs\msbuild-20260630_215517_208.log` and
  `Z:\src\RedSalamander\.build\logs\msbuild-20260630_220028_025.log`.

Important root-cause notes:

- The FileOps crash dump resolved to stack exhaustion during
  `FileOperationState::StartOperation -> EnsurePopupVisible ->
  FileOperationsPopupState::WndProc/OnNcPaint/Render/EnsureTarget`, before
  `Task::ThreadMain` logged `task.started`. The dump had `RSP` 32 bytes below
  the captured stack base and about 1 MB captured stack.
- The focused Floodgate failure after the crash fix was selftest-only path
  mismatch: the hook was armed with `generic_wstring()` slash style while the
  bridge passed a native path.

Current blocker:

```text
No known Commands/FileOps crash blocker remains.
Remaining closeout gates are downstream verification and cleanup.
```

Temporary diagnostics to remove before final verification:

- `RedSalamander\IconCache.cpp`: `IconCache::ConvertIconToBitmap` debug
  milestones.
- `RedSalamander\FolderView.Icons.cpp`: `FolderView::OnCreateIconBitmap`
  selftest trace milestones.

Recommended next commands:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2
.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -FailFast -TimeoutMultiplier 2
.\.build\x64\Debug\ViewerPETests.exe
Invoke-Pester -Path .\Tools\Tests -PassThru
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild
```

Then rerun the final Debug and test-enabled Release FolderView perf matrix from
current HEAD, archive it, update the WIP plan findings, migrate any durable
rules, and move the plan to `Specs\Plans\Done\` only after every gate is green.

## 2026-06-30 15:40 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-06-30_154019_folderview_warpdrive_pause\
```

Resume from this archive, not the 14:55 archive. It includes the current tracked
patch, git status/name-status/stat, untracked list, plan/baton snapshots, copied
small `.build\codex-runs` evidence, latest clean Debug build log,
`CONTINUATION.md`, manifest, and SHA256 checksums. It intentionally does not
copy volatile `last_run` recursively.

New evidence since 14:55:

- Debug build passed after the Pack prompt lifecycle diagnostics:
  `Z:\src\RedSalamander\.build\logs\msbuild-20260630_150318_443.log`
  (0 warnings / 0 errors).
- The explicit Pack-prompt suffix is now green:
  `.build\codex-runs\commands_pack_prompt_suffix_628_643_lifecycle_trace2_20260630`
  (16 passed / 0 failed / 0 skipped).
- The earlier Pack prompt crash did not reproduce with the added lifecycle
  traces. The trace shows normal confirm and modal-loop unwind, including
  `ShowModal loop exit` and `ShowModal end`.
- Broad Commands now gets past Pack prompt and fails later:
  `.build\codex-runs\commands_broad_after_pack_lifecycle_trace2_20260630`
  (551 passed / 1 failed / 224 skipped).

Current blocker:

```text
cmd_pane_fileops_speedLimit_prompt_long_run_open_close_stays_stable
Reason: Custom speed-limit prompt ValuePattern should start with '64,0 KB' during cycle 2.
```

Root-cause note:

- Treat this as the active red point. The old Pack prompt crash remains useful
  context, but it is no longer the first blocker.
- Inspect `TestFileOperationsSpeedLimitPromptLongRunOpenCloseStaysStable(...)`,
  `CollectVisibleDescendantValuePatternState(...)`, and the
  `FileOperationsSpeedLimitPromptWindow` snapshot/debug APIs.
- If focused repro is unclear, add diagnostics for cycle number,
  `snapshot.text`, observed `valueState->name`, observed `valueState->value`,
  read-only state, visible edit count, ValuePattern candidate count, prompt
  HWND, and focus/foreground HWNDs.
- Do not treat the UIA helper, prompt provider, or test wait as root cause
  until the observed value candidates are captured.

Recommended next commands:

```powershell
Get-Content RedSalamander\SelfTest\Commands\Commands.SelfTest.FileOps.cpp | Select-Object -Skip 5710 -First 560
rg -n "CollectVisibleDescendantValuePatternState|UiaValuePatternState|ValuePattern" RedSalamander\SelfTest Common -g "*.cpp" -g "*.h"
Get-Content RedSalamander\FolderWindow.FileOperations.Popup.cpp | Select-Object -Skip 873 -First 520
Get-Content RedSalamander\FolderWindow.FileOperations.Popup.cpp | Select-Object -Skip 1190 -First 150
Get-Content RedSalamander\FolderWindow.FileOperations.Popup.cpp | Select-Object -Skip 8080 -First 70
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_pane_fileops_speedLimit_prompt_long_run_open_close_stays_stable
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_pane_fileops_speedLimit_prompt_uses_dxui_surface,cmd_pane_fileops_speedLimit_prompt_live_dx_interaction,cmd_pane_fileops_speedLimit_prompt_long_run_open_close_stays_stable
```

After this blocker is fixed, continue with broad Commands, FileOperations,
ViewerPETests, Tools Pester, Full, final FolderView perf evidence, and plan
move-to-Done only after every gate is green.

## 2026-06-30 14:30 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-06-30_143026_folderview_warpdrive_pause\
```

Resume from this archive, not the 14:05 archive. It includes the current tracked
patch, git status/name-status/stat, untracked list, plan/baton snapshots,
copied `.build\codex-runs` evidence, copied Debug build log, WER crash events,
crash-dump manifest, `CONTINUATION.md`, parsed run summary JSON, and SHA256
checksums.

New evidence since 14:05:

- Debug build passed after the Compare Options helper/diagnostic patch:
  `Z:\src\RedSalamander\.build\logs\msbuild-20260630_140542_081.log`
  (0 warnings / 0 errors).
- The 14:05 Compare Options theme-cycle blocker is green:
  `.build\codex-runs\commands_compare_options_theme_focused_valuepattern_fix_20260630`
  (1 passed / 0 failed / 0 skipped) and
  `.build\codex-runs\commands_compare_options_cluster_valuepattern_fix_20260630`
  (10 passed / 0 failed / 0 skipped).
- Root cause fixed there: `CollectVisibleDescendantValuePatternState(...)`
  now scans all visible edit candidates until one exposes `UIA_ValuePatternId`;
  previously it failed on the first visible edit without ValuePattern even when
  a later visible edit had it.
- Broad Commands now gets past Compare Options but crashes with
  `-1073741819` (`0xC0000005`) at the Pack prompt case:
  `.build\codex-runs\commands_broad_after_valuepattern_fix_20260630`.
- Focused pack prompt and the immediate 3-case cluster pass:
  `.build\codex-runs\commands_pack_prompt_focused_after_valuepattern_fix_20260630`
  (1 passed / 0 failed / 0 skipped) and
  `.build\codex-runs\commands_pack_prompt_three_case_cluster_20260630`
  (3 passed / 0 failed / 0 skipped).
- Explicit 44-case suffix reproduces the native crash:
  `.build\codex-runs\commands_pack_prompt_suffix_600_643_names_20260630`.
- WER identifies both fresh crashes as `textinputframework.dll` access
  violations: exception `0xc0000005`, fault offset `0x000000000006a419`.

Current blocker:

```text
cmd_pane_pack_prompt_uses_dxui_unique_archive_and_packer_extensions
Only fails after the broad/suffix predecessor sequence.
Focused and immediate 3-case cluster runs pass.
Crash: textinputframework.dll 0xc0000005 at offset 0x6a419.
```

Root-cause note:

- Treat this as a DxUi native text input / TSF lifetime or order-residue crash
  until proven otherwise.
- Inspect `Common\DxUi\DxUi.NativeTextInput.cpp`,
  `Common\DxUi\DxUi.TextStoreACP.cpp`,
  `Common\DxUi\DxUi.WindowHost.cpp`, and
  `RedSalamander\FolderWindow.FileSystem.Commands.Part.cpp`.
- Add narrow lifecycle tracing around pack prompt creation, native text input
  activation/deactivation, and text-store teardown before patching behavior.
- Generate suffix repro filters from explicit case names; numeric range syntax
  such as `600..643` is treated as a literal case name by this runner path.

Recommended next commands:

```powershell
Get-Content Common\DxUi\DxUi.NativeTextInput.cpp | Select-Object -Skip 580 -First 170
Get-Content Common\DxUi\DxUi.TextStoreACP.cpp | Select-Object -Skip 1080 -First 100
Get-Content Common\DxUi\DxUi.WindowHost.cpp | Select-Object -Skip 1130 -First 60
Get-Content RedSalamander\FolderWindow.FileSystem.Commands.Part.cpp | Select-Object -Skip 4500 -First 220
Get-Content RedSalamander\FolderWindow.FileSystem.Commands.Part.cpp | Select-Object -Skip 4760 -First 220
Get-Content RedSalamander\FolderWindow.FileSystem.Commands.Part.cpp | Select-Object -Skip 4980 -First 180
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_pane_pack_prompt_uses_dxui_unique_archive_and_packer_extensions
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_pane_changeAttributes_options_dialog_uses_dxui_not_win32_template,cmd_pane_makeFileList_options_dialog_uses_dxui_not_win32_template,cmd_pane_pack_prompt_uses_dxui_unique_archive_and_packer_extensions
```

After broad Commands is green, continue with FileOperations, ViewerPETests,
Tools Pester, Full, final FolderView perf evidence, and plan move-to-Done only
after every gate is green.

## 2026-06-30 14:05 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-06-30_140500_folderview_warpdrive_pause\
```

Resume from this archive, not the 13:01 archive. It includes the current tracked
patch, git status/name-status/stat, untracked list, plan/baton snapshots, copied
focused/cluster/broad `.build\codex-runs` evidence, copied Debug build logs,
`CONTINUATION.md`, parsed run summary JSON, and SHA256 checksums.

New evidence since 13:01:

- Debug build passed after Preferences reverse-navigation diagnostics:
  `Z:\src\RedSalamander\.build\logs\msbuild-20260630_131339_002.log`
  (0 warnings / 0 errors).
- Preferences reverse-navigation is green in focused and predecessor-cluster
  reruns:
  `.build\codex-runs\commands_prefs_category_reverse_focused_diag2_20260630`
  (1 passed / 0 failed / 0 skipped),
  `.build\codex-runs\commands_prefs_compare_to_reverse_cluster_diag2_retry_20260630`
  (12 passed / 0 failed / 0 skipped), and
  `.build\codex-runs\commands_prefs_350_to_reverse_cluster_diag2_retry_20260630`
  (33 passed / 0 failed / 0 skipped).
- Broad Commands moved past Preferences and failed at Shortcuts Escape:
  `.build\codex-runs\commands_broad_after_reverse_diag2_20260630`
  (496 passed / 1 failed / 279 skipped).
- Debug build passed after Shortcuts Escape diagnostics:
  `Z:\src\RedSalamander\.build\logs\msbuild-20260630_133552_412.log`
  (0 warnings / 0 errors).
- Shortcuts Escape is green in focused and block reruns:
  `.build\codex-runs\commands_shortcuts_escape_focused_diag3_20260630`
  (1 passed / 0 failed / 0 skipped) and
  `.build\codex-runs\commands_shortcuts_block_to_escape_diag3_20260630`
  (12 passed / 0 failed / 0 skipped).
- Broad Commands now fails later:
  `.build\codex-runs\commands_broad_after_shortcuts_diag3_20260630`
  (525 passed / 1 failed / 250 skipped).

Current blocker:

```text
cmd_compare_directories_options_theme_cycle_keeps_surface_legible
Reason: Compare Directories options visible DX edit disappeared after the rainbow theme update.
```

Root-cause note:

- Treat this as order-sensitive until proven otherwise.
- The failure is in the Compare Directories Options theme-cycle test after the
  rainbow theme update, where the test expects a visible DX edit to still be
  discoverable through UIA.
- Add diagnostic context in
  `RedSalamander\SelfTest\Commands\Commands.SelfTest.CompareOptions.cpp` before
  changing production behavior: all visible ValuePattern states, expected edit
  identity/value, options snapshot details, focus/active/foreground HWNDs, and
  the UIA state after the theme repaint.

Recommended next commands:

```powershell
Get-Content -Path RedSalamander\SelfTest\Commands\Commands.SelfTest.CompareOptions.cpp | Select-Object -Skip 1480 -First 230
Get-Content -Path RedSalamander\SelfTest\Commands\Commands.SelfTest.CompareOptions.cpp | Select-Object -First 180
rg -n "CollectVisibleDescendantControlValueState|CollectVisibleDescendantControlValueStates|ControlValueState|describeOptionsSnapshot|waitForOptionsSnapshot|CompareDirectoriesOptions" RedSalamander\SelfTest\Commands\Commands.SelfTest.CompareOptions.cpp RedSalamander\SelfTest\Commands\Commands.SelfTest.Settings.cpp RedSalamander\CompareDirectoriesWindow.cpp RedSalamander\CompareDirectoriesWindow.Options.cpp
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_compare_directories_options_theme_cycle_keeps_surface_legible
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_compare_directories_options_uses_dxui_labels_without_visible_legacy_statics,cmd_compare_directories_window_uses_dxui_menu_bar_and_banner_buttons,cmd_compare_directories_options_long_run_open_close_stays_stable,cmd_compare_directories_options_live_dx_body_interaction,cmd_compare_directories_options_pointer_click_toggles_live_dx_interaction,cmd_compare_directories_options_tab_traversal_live_dx_interaction,cmd_compare_directories_options_scroll_to_lower_cards_stays_stable,cmd_compare_directories_options_enter_and_escape_route_default_cancel,cmd_compare_directories_options_access_keys_focus_expected_controls,cmd_compare_directories_options_theme_cycle_keeps_surface_legible
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2
```

After broad Commands is green, continue with FileOperations, ViewerPETests,
Tools Pester, Full, final FolderView perf evidence, and plan move-to-Done only
after every gate is green.
