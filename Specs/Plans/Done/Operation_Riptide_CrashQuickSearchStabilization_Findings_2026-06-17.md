# Operation Riptide - Crash and Quick Search Stabilization Findings

**Status:** Done - superseded historical checkpoint.
**Date:** 2026-06-17
**Branch/worktree:** `claude/zealous-poitras-ce383a` at `Z:\src\RedSalamander\.claude\worktrees\zealous-poitras-ce383a`
**Purpose:** Archive the current crash/focus investigation so work can continue later without losing context.

## Closeout Update - 2026-06-19

This checkpoint is no longer an active blocker. The downstream Quick Search same-folder refresh/focus fix, temporary diagnostic cleanup, authoritative spec updates, and final green Full gate are closed in:

- `Specs\Plans\Done\Operation_Floodgate_RiptideMergeCloudDataSafetyAndCloseout_2026-06-17.md`
- `Specs\Plans\Done\Operation_Riptide_FairstreamRemediation_DataSafetyConflictParity_2026-06-15.md`

The historical notes below are retained for crash/focus triage context only. Any "do not merge", dirty-worktree, temporary-diagnostic, or current-blocker wording below describes the 2026-06-17 checkpoint and is superseded by the final 2026-06-19 closeout records.

## Historical Executive State

Do not merge this branch to `master` yet. The original Access Violation crash path has targeted green evidence, but the full Commands selftest is still not green. The current blocker is `cmd_pane_quickSearch_integrated_navigation` in the full fail-fast run after the Find dialog cases.

The worktree is dirty and contains the broader Riptide/Fairstream remediation changes plus the crash/focus investigation edits. There are also temporary selftest-only diagnostics that must either drive the next fix or be removed before final merge.

## Verified Evidence So Far

### Build

Latest build after adding the current selftest diagnostics:

```powershell
& '.\build.ps1' -ProjectName RedSalamander -Configuration Debug
```

Result: success, `0 warning(s), 0 error(s)`.

Log:

```text
.build\logs\msbuild-20260617_195352_165.log
```

### Crash Dumps

No new crash dumps were created during the latest rebuilt test runs. Latest crash artifacts remain old:

```text
C:\Users\eric\AppData\Local\RedSalamander\Crashes\RedSalamander-20260617-134447-p85180.dmp
C:\Users\eric\AppData\Local\RedSalamander\Crashes\RedSalamander-20260617-134447-p85180.txt
```

The user-provided dump path was:

```text
C:\Users\eric\AppData\Local\RedSalamander\Crashes\RedSalamander-20260617-130846-p80744.dmp
```

Both analyzed crash reports pointed at the same shutdown/destruction class around `DxUi::WindowHost::Detach` and command-window teardown.

### Targeted Passes

These passed in visible GUI selftest runs after fixes:

```powershell
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-fail-fast --selftest-case=cmd_connection_manager_window_escape_from_dx_input_closes_cancel
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-fail-fast --selftest-case=pane_view_options_toggle_preview_pane_tabs_and_selection
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-fail-fast --selftest-case=cmd_plugin_configuration_dialog_tab_traversal_live_dx_interaction
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-fail-fast --selftest-case=cmd_preferences_dialog_keyboard_reordered_resized_columns_survive_search_roundtrip
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-fail-fast --selftest-case=cmd_pane_quickSearch_integrated_navigation
```

Also passed:

```powershell
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-fail-fast --selftest-case=pane_
```

Archived pass evidence includes:

```text
Specs\TestRuns\4cb089111a23\Commands\2026-06-17_185625
Specs\TestRuns\4cb089111a23\Commands\2026-06-17_192517
Specs\TestRuns\4cb089111a23\Commands\2026-06-17_194354
```

## Fixed or Hardened During This Investigation

### 1. Connection Manager Escape UAF / Access Violation

Finding:

- The fatal user crash hit `DxUi::WindowHost::Detach` through `ConnectionManagerWindow::DebugRouteCommandKey`.
- `DebugRouteCommandKey(VK_ESCAPE)` could route both destructive keydown and later keyup/touch logic after Cancel destroyed the dialog/window implementation.

Current fix:

- `RedSalamander\ConnectionManagerWindow.cpp`
- `DebugRouteCommandKey(VK_ESCAPE)` returns immediately after the destructive keydown route.
- `WM_NCDESTROY` cleanup is covered when DxUi consumes the message.

Evidence:

- Exact crash regression case passed.
- No new dump files after current rebuilt runs.

### 2. Main Menu Host Shutdown Crash Class

Finding:

- The earlier pasted stack showed `MainMenuBarHost` destruction calling into `DxUi::WindowHost::Detach` during process exit.
- The host/window ownership ordering was unsafe around `wil::unique_hwnd` teardown and `WM_NCDESTROY`.

Current fix:

- `RedSalamander\RedSalamander.cpp`
- `MainMenuBarHost::~MainMenuBarHost()` calls `Destroy()`.
- `Destroy()` detaches host state and clears userdata before `_hwnd.reset()`.
- `OnNcDestroy()` detaches and releases `_hwnd` without double-owning a destroyed HWND.
- `_host` member now destructs before `_hwnd` owner.

Risk:

- This is the right shape for the observed shutdown stack, but full-suite verification is still blocked by Quick Search before final merge.

### 3. Preferences Category Host Focus Flake

Finding:

- Full Commands fail-fast hit:

```text
cmd_preferences_dialog_keyboard_reordered_resized_columns_survive_search_roundtrip
Failed to focus the Preferences category host for Keyboard reordered-resized/search validation.
```

- The direct `SetFocus(categoryTreeHost) == categoryTreeHost` assertion was brittle after heavy visible UI test sequencing.

Current fix:

- `RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp`
- Replaced the raw one-shot focus assertion for the failing case with existing `FocusWindowAndWait(...)` plus diagnostics.

Evidence:

- Exact case passed after the change.
- Later full fail-fast progressed past this case and reached Quick Search.

### 4. Quick Search Focus/Activation Hardening

Finding:

- Exact `cmd_pane_quickSearch_integrated_navigation` passes.
- Full fail-fast fails only after preceding Find dialog cases.
- Earlier failures alternated between initial command activation and later reactivation/focus holes.

Current hardening:

- `RedSalamander\FolderWindow.FileSystem.Navigation.Part.cpp`
  - `CommandQuickSearch(...)` refocuses pane after activation and posts `WndMsg::kPaneRestoreFolderFocus` to the root owner.
- `RedSalamander\FolderView.Interaction.cpp`
  - `OnKillFocusMessage(HWND newFocus)` preserves Quick Search when `WM_KILLFOCUS` has a null target.
  - If null-focus happens while Quick Search is active, posts a guarded folder-focus restore.

Evidence:

- Exact Quick Search case passed.
- Full fail-fast now shows `WM_COMMAND` reaches Quick Search and activation succeeds, but printable typing still loses active state in the contaminated sequence.

## Current Blocker

Latest full fail-fast run before adding the final exit/focus trace:

```text
Specs\TestRuns\4cb089111a23\Commands\2026-06-17_195056
```

Failure:

```text
cmd_pane_quickSearch_integrated_navigation
Quick Search initial typing should advance to query 'al'; active=0, query='', focused='alpha.txt'; focus=0x12021FB8, focusedFolderView=0x12021FB8, focusedPane=Left, expectedLeftFolderView=0x12021FB8.
```

Important trace from `selftest_run_trace.txt`:

```text
WM_COMMAND QuickSearch: pane=left focus=0x12021FB8 focusedView=0x12021FB8
CommandQuickSearch: begin pane=left
CommandQuickSearch: after activate active=1 query='' focused='gamma.txt' focus=0x12021FB8
FolderView::SetNameFilterState: enabled=0 trimmed='' refresh=0
Case failure: cmd_pane_quickSearch_integrated_navigation reason='Quick Search initial typing should advance to query 'al'; active=0, query='', focused='alpha.txt'; focus=0x12021FB8, focusedFolderView=0x12021FB8, focusedPane=Left, expectedLeftFolderView=0x12021FB8.'
```

Conclusion:

- This is no longer a missed `WM_COMMAND`.
- It is no longer simply "wrong focused pane".
- `CommandQuickSearch` sets `active=1`.
- After printable typing, focus remains on the same folder view, item focus moves to `alpha.txt`, but Quick Search has been cleared.
- The likely culprit is a state transition after first/second printable char, possibly through `OnCharMessage`, `FocusItem`, `NotifyFocusedItemChanged`, `NotifyIncrementalSearchChanged`, a focus-change message, or an enumeration/path/filter side effect.

## Temporary Diagnostics Currently In Code

These were added to continue root cause. Keep them for the next run, then remove or demote once the bug is fixed.

### `RedSalamander\RedSalamander.cpp`

- Selftest trace in `IDM_PANE_QUICK_SEARCH` showing focused pane and focused folder view before dispatch.

### `RedSalamander\FolderWindow.FileSystem.Navigation.Part.cpp`

- Selftest trace at `CommandQuickSearch` begin.
- Selftest trace immediately after `ActivateIncrementalSearch`.

### `RedSalamander\FolderView.Interaction.cpp`

Added after the latest full fail-fast run and rebuilt successfully, but not yet used in a full rerun:

- Selftest trace in `OnKillFocusMessage(HWND newFocus)`.
- Selftest trace in `ExitIncrementalSearch()`.
- `DescribeQuickSearchFocusTarget(...)` helper under `ENABLE_TESTS`.

Next full run should show exactly whether the post-activation clear is coming from focus loss, key handling, enumeration, or another explicit `ExitIncrementalSearch`.

## Next Runbook

### 1. Reproduce with the new exit/focus trace

Run visibly, not hidden:

```powershell
$p = Start-Process -FilePath '.\.build\x64\Debug\RedSalamander.exe' -ArgumentList @('--commands-selftest','--selftest-fail-fast') -PassThru
if (-not $p.WaitForExit(1800000)) { Stop-Process -Id $p.Id -Force; "TimedOutKilled Id=$($p.Id)" } else { "ExitCode=$($p.ExitCode)" }
```

Then inspect:

```powershell
Get-ChildItem 'Specs\TestRuns\4cb089111a23\Commands' | Sort-Object LastWriteTime -Descending | Select-Object -First 1
Select-String -Path '<latest>\selftest_run_trace.txt' -Pattern 'QuickSearch|CommandQuickSearch|OnKillFocus|QuickSearch Exit|Case failure' -Context 0,4
Get-ChildItem 'C:\Users\eric\AppData\Local\RedSalamander\Crashes' | Sort-Object LastWriteTime -Descending | Select-Object -First 5 Name,LastWriteTime,Length
Get-Process RedSalamander -ErrorAction SilentlyContinue
```

Expected next evidence:

- If `QuickSearch Exit` appears, fix the caller path that exits during printable typing.
- If `OnKillFocus` has a non-null target before the clear, decide whether that target is an in-pane child that should not cancel Quick Search.
- If no exit trace appears but active becomes false, inspect direct state mutation of `_incrementalSearch.active`.

### 2. Inspect likely code paths

Primary files:

```text
RedSalamander\FolderView.Interaction.cpp
RedSalamander\FolderView.Selection.cpp
RedSalamander\FolderView.Enumeration.cpp
RedSalamander\FolderView.cpp
RedSalamander\FolderWindow.StatusBar.cpp
RedSalamander\FolderWindow.FileSystem.Navigation.Part.cpp
RedSalamander\SelfTest\Commands\Commands.SelfTest.Search.cpp
```

Key functions:

```text
FolderView::OnCharMessage
FolderView::FocusItem
FolderView::ExitIncrementalSearch
FolderView::OnKillFocusMessage
FolderView::ProcessEnumerationResult
FolderWindow::UpdatePaneStatusBar
FolderWindow::SetFolderPath
FolderWindow::CommandQuickSearch
TestPaneQuickSearchIntegratedNavigation
```

### 3. Fix rule

Do not paper over the Quick Search test with retries. The full-run failure proves a real cross-case state leak or focus/state transition. The product contract should be:

- Quick Search command activates the focused pane.
- Printable chars keep Quick Search active while focus remains in the folder view.
- Transient null focus must not cancel Quick Search.
- Moving focus to a real external control may cancel Quick Search, but in-pane status/filter/history synchronization must not.

### 4. Verification before merge

Minimum before considering commit/merge:

```powershell
& '.\build.ps1' -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-fail-fast --selftest-case=cmd_connection_manager_window_escape_from_dx_input_closes_cancel
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-fail-fast --selftest-case=cmd_preferences_dialog_keyboard_reordered_resized_columns_survive_search_roundtrip
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-fail-fast --selftest-case=cmd_pane_quickSearch_integrated_navigation
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-fail-fast
```

After every GUI run:

```powershell
Get-ChildItem 'C:\Users\eric\AppData\Local\RedSalamander\Crashes' | Sort-Object LastWriteTime -Descending | Select-Object -First 5 Name,LastWriteTime,Length
Get-Process RedSalamander -ErrorAction SilentlyContinue
```

Do not run UI selftests hidden. Hidden top-level windows can invalidate focus/visibility assertions.

## Current Dirty Worktree Notes

Branch status at archive time:

```text
claude/zealous-poitras-ce383a...origin/claude/zealous-poitras-ce383a [ahead 58]
```

High-signal touched files from this investigation:

```text
RedSalamander\ConnectionManagerWindow.cpp
RedSalamander\RedSalamander.cpp
RedSalamander\FolderWindow.FileSystem.Navigation.Part.cpp
RedSalamander\FolderView.Interaction.cpp
RedSalamander\ManagePluginsDialog.cpp
RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp
RedSalamander\SelfTest\Commands\Commands.SelfTest.Search.cpp
RedSalamander\SelfTest\Commands\Commands.SelfTest.Settings.cpp
Common\DxUi\DxUi.WindowHost.cpp
```

There are many other dirty files from the broader Riptide/Fairstream branch. Do not revert them casually. Review diffs before staging.

Untracked at archive time:

```text
Specs\TestRuns\SINON\
```

There is also a separate running process from the main workspace, not from this worktree:

```text
Z:\src\RedSalamander\.build\x64\Release\RedSalamander.exe
```

Leave it alone unless the user explicitly asks to close it.

## Merge Status

Merge to `master` has not been done and should not be done yet. The final instruction remains:

1. Fix the remaining blocker.
2. Remove temporary diagnostics or intentionally keep only useful selftest diagnostics.
3. Run required verification.
4. Move/close the appropriate WIP spec only after green evidence.
5. Commit and merge to `master`.

