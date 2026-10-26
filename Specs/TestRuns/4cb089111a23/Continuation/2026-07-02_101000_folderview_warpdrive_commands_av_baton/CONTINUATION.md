# FolderView Warp Drive Continuation Baton

Archive created: 2026-07-02 10:10 +02:00

Branch: `codex/folderview-warpdrive`
HEAD: `275c040348a3994d45c6aef340333a35a0490cb6`
Origin state: `0 25` (`origin/master...HEAD`)

## Current Status

Do not move `Specs\Plans\WIP\Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md` to Done yet.

The latest blocker is broad Commands reliability, not a FolderView perf budget failure. The previous suspected Find-dialog stop did not reproduce in focused, suffix, 45-case, or exact 456-case prefix runs. A later full broad Commands rerun progressed through the FolderView perf block and then crashed with access violation before normal Commands results were written.

## Evidence Captured

- `DxUiTests` Debug build passed after the MenuBar whitespace-insensitive source guard fix: `run-artifacts\dxuitests_build_after_menubar_guard_warningfix_20260702.log`, MSBuild log `run-artifacts\msbuild-20260702_084003_855.log`.
- `DxUiTests.exe` passed: `run-artifacts\dxuitests_after_menubar_guard_warningfix_20260702.log`.
- The exact 96-case Commands slice from the prior focus/prompt-close baton passed again: `run-artifacts\commands_slice_532_627_repeat_after_dxuitests_20260702.log`.
- First broad Commands attempt timed out/hung near `cmd_pane_find_dialog_restored_combined_view_state_action_buttons_activate_selection`; captured under `run-artifacts\commands_broad_hang_find_action_buttons_20260702\`.
- Focused Find action-buttons, suffix-5, suffix-15, suffix-45, and exact prefix-456 reruns all passed. The prefix run is archived as `run-artifacts\commands_prefix456_to_find_action_buttons_20260702.log` with exit file `run-artifacts\commands_prefix456_to_find_action_buttons_20260702.exit.txt` (`0`).
- Second broad Commands rerun exited `1`; its selftest child exited `-1073741819` (`0xC0000005`). Last trace reaches 775 cases and stops at `cmd_pane_selection_invert`. Evidence:
  - `run-artifacts\commands_broad_rerun_after_prefix456_green_20260702.log`
  - `run-artifacts\commands_broad_rerun_after_prefix456_green_20260702.exit.txt`
  - `run-artifacts\commands-broad-av-run-all-tests-results.json`
  - `run-artifacts\commands-broad-av-trace.txt`
  - `run-artifacts\commands-broad-av-case-order.txt`
- WER reports `textinputframework.dll`, exception `0xc0000005`, fault offset `0x000000000006a419`; event log copy is `windows-error-reporting-recent.txt`.
- Latest local crash dump was not copied into the repo because it is about 120 MB. Pointer is in `latest-crash-dump-pointer.txt`: `C:\Users\eric\AppData\Local\CrashDumps\RedSalamander.exe.73392.dmp`.
- Current staged diff snapshot: `patches\staged-diff.patch`.
- Current unstaged focus/prompt/DxUi source-guard diff snapshot: `patches\current-unstaged-focus-prompt-dxui.patch`.
- Current WIP plan snapshot: `Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.snapshot.md`.

## Active Code Changes Since The Previous Baton

- `Tests\DxUiTests\DxUiTests.Controls.cpp`: MenuBar paint source guard now normalizes ASCII whitespace before checking for cached layout usage, fixing a brittle test-only guard against aligned code.
- `RedSalamander\SelfTest\Commands\Commands.SelfTest.Settings.cpp`: `FocusFolderViewPane(...)` waits for stable Win32 focus, focused FolderView HWND, and focused pane state while pumping messages.
- `RedSalamander\FolderWindow.FileOperations.Popup.cpp`: File Operations speed-limit prompt debug confirm/cancel paths now post a close request instead of synchronously sending `WM_CLOSE`.
- `Tests\DxUiTests\DxUiTests.NativeTextInput.cpp`: source guard covers File Operations speed-limit prompt post-close behavior.
- `RedSalamander\FolderWindow.FileSystem.Navigation.Part.cpp`: prior FolderView prompt post-close fix remains in the active unstaged diff.

## Next Move

Start from the crash boundary, not from broad Commands.

1. Confirm no stale app process is running:
   ```powershell
   Get-Process RedSalamander -ErrorAction SilentlyContinue
   ```
2. Run the focused boundary:
   ```powershell
   .\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -CaseFilter 'cmd_pane_selection_invert' -TimeoutMultiplier 2.0
   ```
3. Run the archived last-10 and last-20 filters:
   ```powershell
   $filter = (Get-Content Specs\TestRuns\4cb089111a23\Continuation\2026-07-02_101000_folderview_warpdrive_commands_av_baton\run-artifacts\commands-broad-av-last10.filter.txt) -join ','
   .\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -CaseFilter $filter -TimeoutMultiplier 2.0

   $filter = (Get-Content Specs\TestRuns\4cb089111a23\Continuation\2026-07-02_101000_folderview_warpdrive_commands_av_baton\run-artifacts\commands-broad-av-last20.filter.txt) -join ','
   .\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -CaseFilter $filter -TimeoutMultiplier 2.0
   ```
4. If reduced filters are green, build a prefix ending at `cmd_pane_selection_invert` from `run-artifacts\commands-broad-av-case-order.txt` and run that once with `-FailFast`.
5. If a slice reproduces the crash, copy `last_run`, WER, and the crash-dump pointer before any next run. Only then inspect the dump/TSF stack and patch.

## Remaining Before Done

- Broad Commands must be green on the current worktree.
- Full suite must pass.
- Final Debug and test-enabled Release FolderView perf matrices must be rerun from current HEAD and cited.
- Durable closeout notes must be migrated to authoritative specs.
- Then, and only then, move the WIP plan to `Specs\Plans\Done\`.
