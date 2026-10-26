# FolderView WarpDrive continuation baton - Commands race checkpoint

Created: 2026-07-03 10:59 local
Repo: `Z:\src\RedSalamander`
Branch: `codex/folderview-warpdrive`
HEAD: `275c040348a3994d45c6aef340333a35a0490cb6`
Origin master captured in `git-branch-head.txt`.

## Current state

This baton archives the state after continuing convergence on the FolderView WarpDrive plan, specifically the post-Full red Commands-suite fallout.

Do not move `Specs/Plans/WIP/Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md` to Done yet. The plan is still open until:

- broad Commands is green after the remaining order-sensitive races are fixed,
- ViewerPE direct is rechecked,
- Monitor ETW latency direct is rechecked,
- Full suite is green,
- final Debug and test-enabled Release FolderView perf matrices are archived and cited,
- the WIP plan is updated with final evidence and moved to `Specs/Plans/Done/`.

## What changed since the previous baton

### Fixed and verified

`cmd_pane_find_dialog_action_buttons_activate_expected_commands` failed in broad Commands because the test selected the directory result immediately after the Find Files window reported idle. Idle was not enough: the deferred results grid could still be behind.

Patch:

- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Search.cpp`
- Added a bounded `WaitForFindSnapshot(...)` that waits until the directory result exposes exactly the expected `sub` path before calling `DebugSelectFindFilesWindowResult(...)`.

Evidence:

- Rebuild: `run-artifacts/msbuild_redsalamander_after_find_action_buttons_wait_20260703.log`
  - `0 warning(s), 0 error(s)`
- Focused case: `run-artifacts/runall_commands_find_action_buttons_focused_after_wait_20260703.run-all.json`
  - `1 passed / 0 failed`
- Find-dialog prefix: `run-artifacts/runall_commands_find_dialog_prefix_after_wait_20260703.run-all.json`
  - `63 passed / 0 failed`

### Broad Commands result after that fix

Evidence:

- `run-artifacts/commands_broad_after_find_action_buttons_wait_20260703.run-all.json`
- `run-artifacts/commands_broad_after_find_action_buttons_wait_20260703.trace.txt`

Result:

- `778 total / 774 passed / 2 failed / 2 skipped`
- Duration: about `1044089 ms`
- The previous Find action-buttons failure is gone.

New failures:

1. `cmd_pane_navigationView_path_doubleClick_after_focusAddressBar_tab_traversal`
   - Failure: `Focused window unavailable before navigation-view path edit tab handoff.`
   - Focused rerun passed.
   - Likely root cause: selftest race. The test calls one-shot `GetFocus()` immediately after path edit-mode setup, under a broad-suite ordering where focus can lag behind the edit snapshot.

2. `cmd_app_menuBar_persistent_view_to_plugins_hover_switches_popup`
   - Failure: `Delivering a Plugins hover while View is open did not select the Plugins top-level menu... selectedAfterWait=5, expectedPluginsIndex=4, pointerSwitches=2`
   - Focused rerun passed.
   - Trace shows popup replacement happened and `pointerSwitches=2`, but selected top-level index still sampled as View.
   - Likely root cause: selftest readiness race around menu hover/popup replacement versus selected top-level convergence.

Focused recheck evidence:

- `run-artifacts/runall_commands_new_failures_focused_summary_20260703.txt`
- `run-artifacts/runall_commands_cmd_pane_navigationView_path_doubleClick_after_focusAddressBar_tab_traversal_focused_after_broad_20260703.run-all.json`
- `run-artifacts/runall_commands_cmd_app_menuBar_persistent_view_to_plugins_hover_switches_popup_focused_after_broad_20260703.run-all.json`

Both focused failures passed.

## Next exact work

1. Patch `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`.
   - In `TestPaneNavigationViewPathDoubleClickEntersEditMode`, replace the one-shot `GetFocus()` before the Tab handoff with a bounded wait that confirms focus is on `currentEditInputHwnd` or `currentEditHostHwnd` from `NavigationViewDebugSnapshot`.
   - Keep the patch selftest-only.
   - Add diagnostics that include the last focused hwnd and edit snapshot hwnds.

2. Read and patch the menu case around `TestMainMenuPersistentViewToPluginsHoverSwitchesPopup`.
   - Evidence lines in the trace show the replacement popup happened but selected index lagged.
   - Likely fix is to wait for both conditions before assertion: replacement popup is present and selected top-level index is Plugins.
   - Do not change product code unless the test proves a real product race.

3. Rebuild Debug:

   ```powershell
   .\build.ps1 -ProjectName RedSalamander -Configuration Debug
   ```

4. Rerun focused Commands cases:

   ```powershell
   .\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_pane_navigationView_path_doubleClick_after_focusAddressBar_tab_traversal
   .\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_app_menuBar_persistent_view_to_plugins_hover_switches_popup
   ```

5. Rerun broader Commands:

   ```powershell
   .\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild
   ```

6. After broad Commands is green, recheck:

   ```powershell
   .\.build\x64\Debug\ViewerPETests.exe
   .\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --wait-instance --perf --monitor-etw-burst-mode=latency --monitor-etw-burst-count=60 --monitor-etw-burst-size=260
   .\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild
   ```

7. Only after Full is green, run and archive final FolderView perf matrices in Debug and test-enabled Release, update the WIP plan, then move the plan to Done.

## Archived contents

State:

- `git-status-short.txt`
- `git-branch-head.txt`
- `origin-master-ahead-behind.txt`
- `process-snapshot-redsalamander.txt`
- `crash-sidecars-latest.txt`

Diffs:

- `patches/staged-code-and-spec-no-testruns.patch`
- `patches/unstaged-code-and-spec-no-testruns.patch`
- `git-diff-*-no-testruns.txt`

Run evidence:

- `run-artifacts/commands_broad_after_find_action_buttons_wait_20260703.*`
- `run-artifacts/runall_commands_find_*_after_wait_20260703.*`
- `run-artifacts/runall_commands_new_failures_focused_summary_20260703.txt`
- `run-artifacts/runall_commands_cmd_*_focused_after_broad_20260703.*`
- `run-artifacts/codex-runs-latest-list.txt`

Source snapshots:

- `source-snapshots/RedSalamander_SelfTest_Commands_Commands.SelfTest.Search.cpp`
- `source-snapshots/RedSalamander_SelfTest_Commands_Commands.SelfTest.ViewCommands.cpp`
- `source-snapshots/RedSalamander_SelfTest_CompareDirectories_CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp`
- `source-snapshots/Tools_Tests_ResourceLocalizationContracts.Tests.ps1`
- `source-snapshots/Specs_Plans_WIP_Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md`
- `source-snapshots/.agents_skills_red-salamander-crash-forensics_SKILL.md`

No dump files were copied into this baton.
