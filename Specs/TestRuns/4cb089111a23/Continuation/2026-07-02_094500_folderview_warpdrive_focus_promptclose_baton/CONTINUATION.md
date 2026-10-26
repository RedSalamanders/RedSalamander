# FolderView WarpDrive Continuation Baton - Focus and Prompt Close

Archive created: 2026-07-02
Workspace: `Z:\src\RedSalamander`
Branch: `codex/folderview-warpdrive`
Plan: `Specs/Plans/WIP/Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md`

## Current status

Do not move the plan to Done yet.

The previously pending exact 96-case Commands slice now passes after the latest focus-settle and prompt-close changes:

- `.build\codex-runs\commands_slice_532_627_after_focus_and_fileop_postclose_20260702.log`
- `C:\Users\eric\AppData\Local\RedSalamander\SelfTest\last_run\run-all-tests-results.json`
- Result: Commands Debug, timeout scale 2.0, 96 passed, 0 failed, 0 skipped, exit_code 0.

This proves the reduced TSF/focus slice is green in this run. It does not yet prove the branch is ready to finish. The remaining gates are still broad Commands, FileOps, DxUiTests source guards, Tools/Pester, full suite, and final Debug/test-enabled Release perf evidence.

## Changes since the previous baton

1. `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`
   - Hardened `FocusFolderViewPane(FolderWindow::Pane pane)` from a one-shot `SetFocus(view)` into a bounded condition wait.
   - The helper now pumps messages and requires stable agreement between Win32 focus, `GetFocusedFolderViewHwnd()`, and `GetFocusedPane()` before the command selftest proceeds.
   - Root cause: the exact slice sometimes failed `cmd_pane_selection_select_all_keeps_navigation_shell_stable` even though command semantics were correct (`selectedCount=3`, expected 3, no shell UI left open). The failing evidence showed stale focus bookkeeping only.

2. `RedSalamander/FolderWindow.FileOperations.Popup.cpp`
   - Added `PostFileOperationsSpeedLimitPromptCloseDebugCommand(...)`.
   - Changed `DebugConfirmFileOperationsSpeedLimitPrompt()` and `DebugCancelFileOperationsSpeedLimitPrompt()` to post the close command instead of sending synchronously.
   - Rationale: this speed-limit prompt has the same DxUi/TSF modal teardown shape as the FolderView prompt close path. Posting avoids destroying the prompt inside a cross-thread `SendMessageW` stack.

3. `Tests/DxUiTests/DxUiTests.NativeTextInput.cpp`
   - Extended the prompt-close source guard to cover `FolderWindow.FileOperations.Popup.cpp`.
   - The guard now requires the FileOperations speed-limit prompt debug confirm/cancel helpers to use the post-close helper and to avoid direct `SendMessageW`.

The previous baton already captured the FolderView prompt-close patch in `RedSalamander/FolderWindow.FileSystem.Navigation.Part.cpp`; that remains part of the active fix.

## Evidence archived here

Build logs:

- `run-artifacts\msbuild-20260702_082649_559.log` - Debug RedSalamander build after focus helper, passed with 0 warnings / 0 errors.
- `run-artifacts\msbuild-20260702_083130_192.log` - Debug RedSalamander build after FileOperations post-close/source guard, passed with 0 warnings / 0 errors.

Focused and reduced Commands evidence:

- `run-artifacts\focused_select_all_after_focushelper_20260702.log` - focused Select All case passed.
- `run-artifacts\commands_slice_532_627_after_focushelper_20260702.log` - exact 96-case reduced slice passed after focus helper.
- `run-artifacts\speedlimit_fourcase_after_fileop_postclose_20260702.log` - speed-limit prompt cluster passed after FileOperations post-close helper.
- `run-artifacts\commands_slice_532_627_after_focus_and_fileop_postclose_20260702.log` - exact 96-case reduced slice passed after both focus and FileOperations prompt-close changes.
- `run-artifacts\latest-run-all-tests-results.json` - latest Run-AllTests summary, 96/0/0.
- `run-artifacts\latest-commands-results.json` - latest Commands selftest case results, 96/0/0.

Context snapshots:

- `Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.snapshot.md`
- `git-status-short.txt`
- `git-diff-stat.txt`
- `git-diff-cached-stat.txt`
- `git-diff-name-status.txt`
- `git-diff-cached-name-status.txt`
- `origin-master-ahead-behind.txt`
- `process-snapshot-redsalamander.txt`

Patch snapshots:

- `patches\current-promptclose-focus-sourceguard.patch`
- `patches\staged-diff.patch`

## Next steps

1. Build and run DxUiTests because the native text input source guard changed:
   - `.\build.ps1 -ProjectName DxUiTests -Configuration Debug`
   - `.\.build\x64\Debug\DxUiTests.exe`

2. Rerun the exact 96-case Commands slice once more to check for intermittency:
   - use `run-artifacts\case-filter.txt` or the original `.build\codex-runs\commands_slice_532_627_after_postclose_20260702_0817\case-filter.txt`
   - `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter <filter> -TimeoutMultiplier 2.0`

3. Run broad Commands:
   - `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2.0`

4. Continue with the remaining plan gates:
   - FileOps full coverage
   - ViewerPE/VLC HUD checks
   - Tools Pester tests
   - Full suite
   - final Debug and test-enabled Release FolderView perf matrix

5. Only after all gates are green:
   - update the WIP plan checkboxes and evidence paths
   - merge durable contracts into authoritative specs
   - move the plan to `Specs/Plans/Done/`

## Current caution

The reduced Commands slice is now green, but broad Commands has not been rerun after the latest FileOperations prompt-close/source-guard patch. Treat the branch as promising but still open.
