# FolderView WarpDrive Continuation Baton

Created: 2026-07-03 16:15 Europe/Paris
Workspace: `Z:\src\RedSalamander`
Branch: `codex/folderview-warpdrive`
Branch relation at archive time: `25` ahead / `4` behind `origin/master`

## Current State

This is a pause archive for the active WIP plan:

```text
Specs\Plans\WIP\Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md
```

Do not move the plan to Done yet. The original FolderView perf work is still gated by broad Commands reliability, origin/master resync, proper Full, and the final Debug plus test-enabled Release FolderView perf matrix.

The previous NavigationView full-path popup DPI red is no longer the active edge. It was rechecked after the 15:55 archive and all reduced windows passed:

```text
.build\codex-runs\commands_nav_full_path_popup_predecessor2_20260703.log
Specs\TestRuns\4cb089111a23\Commands\2026-07-03_160054

.build\codex-runs\commands_nav_full_path_popup_predecessor3_20260703.log
Specs\TestRuns\4cb089111a23\Commands\2026-07-03_160114

.build\codex-runs\commands_nav_view_prefix17_recheck_20260703.log
Specs\TestRuns\4cb089111a23\Commands\2026-07-03_160140
```

The current blocker is a fresh broad Commands crash:

```text
.build\codex-runs\commands_broad_after_nav_prefix17_recheck_20260703.log
```

The wrapper exited with `-1073740771` (`0xC000041D`) before normal `commands\results.json` was written. The RedSalamander crash handler produced paired dump/sidecar artifacts in:

```text
C:\Users\eric\AppData\Local\RedSalamander\Crashes
```

Copied crash artifacts:

```text
crash-artifacts\RedSalamander-20260703-161435-p81732.dmp
crash-artifacts\RedSalamander-20260703-161435-p81732.txt
crash-artifacts\last_crash.txt
```

Crash sidecar summary:

```text
ExceptionCode=0xC0000005
ThreadId=75392

00 textinputframework/TSF
11 GetMessageW
12 FileOperationsSpeedLimitPromptWindow::ShowModal
13 ShowCustomSpeedLimitPrompt
14 FileOperationsPopupState::ShowCustomSpeedLimitPromptForTask
15 FileOperationsPopupState::WndProc
19 PumpPendingMessages
20 RunFileOperationsSpeedLimitPromptModalCycle
21 TestFileOperationsSpeedLimitPromptLiveDxInteraction
22 cmd_pane_fileops_speedLimit_prompt_live_dx_interaction
```

The last copied selftest trace stops at:

```text
Case: cmd_pane_fileops_speedLimit_prompt_live_dx_interaction
```

The immediately previous case passed:

```text
cmd_pane_fileops_speedLimit_prompt_uses_dxui_surface
```

## What Changed Since The Previous Baton

- Rechecked the NavigationView full-path popup DPI failure from the 15:55 baton.
- The 2-case predecessor slice passed: `cmd_pane_navigationView_unfocused_pane_click_focuses_target_pane,cmd_pane_navigationView_full_path_popup_edit_route`.
- The 3-case predecessor slice passed: `cmd_pane_navigationView_path_ancestor_click_navigates_to_ancestor,cmd_pane_navigationView_unfocused_pane_click_focuses_target_pane,cmd_pane_navigationView_full_path_popup_edit_route`.
- The exact 17-case NavigationView prefix that had failed earlier now passed: 17 passed / 0 failed / 0 skipped.
- A new broad Commands run then crashed in the File Operations speed-limit prompt modal path before normal results were written.

## Current Root-Cause Lead

Treat this as a new TSF/native input modal-prompt teardown crash until proven otherwise.

The useful code boundary is:

```text
RedSalamander\FolderWindow.FileOperations.Popup.cpp
RedSalamander\SelfTest\Commands\Commands.SelfTest.FileOps.cpp
```

Initial inspection before pause:

- `FileOperationsSpeedLimitPromptWindow::ShowModal()` is in a modal `GetMessageW` loop when TSF crashes.
- `RunFileOperationsSpeedLimitPromptModalCycle(...)` drives the prompt through a selftest worker while the UI thread is inside the prompt modal loop.
- `TestFileOperationsSpeedLimitPromptLiveDxInteraction()` is the active case.
- Do not patch from the stack alone. First inspect whether destructive prompt close/confirm/cancel/debug helpers use worker-thread synchronous UIA invoke or `SendMessageW` into a modal DxUi HWND.

The local reusable crash skill has a matching note for "DxUi Prompt Teardown Crash": prefer posted confirm/cancel/close commands for destructive modal prompt teardown; keep synchronous sends only for non-destructive snapshot/set commands. Verify the concrete implementation before applying that pattern.

## Archive Contents

This archive contains:

- `CONTINUATION.md` - this baton.
- `archive-manifest.txt` - refreshed file list.
- `git-status-short.txt`, `git-status-short-refresh.txt`, `git-head.txt`, `git-branch.txt`, and `origin-master-ahead-behind.txt`.
- `patches\staged-code-and-spec-no-testruns.patch` and `patches\unstaged-code-and-spec-no-testruns.patch`, plus stat/name-status companions.
- `run-artifacts\commands_broad_after_nav_prefix17_recheck_20260703.log`.
- `run-artifacts\commands_nav_full_path_popup_predecessor2_20260703.log`.
- `run-artifacts\commands_nav_full_path_popup_predecessor3_20260703.log`.
- `run-artifacts\commands_nav_view_prefix17_recheck_20260703.log`.
- `commands-archives\2026-07-03_160054`, `2026-07-03_160114`, and `2026-07-03_160140`.
- `last-run\trace.txt` and `last-run\plugin_config_live_step.txt`.
- `source-snapshots\RedSalamander_FolderWindow.FileOperations.Popup.cpp`.
- `source-snapshots\RedSalamander_SelfTest_Commands_Commands.SelfTest.FileOps.cpp`.
- `source-snapshots\RedSalamander_SelfTest_Commands_Commands.SelfTest.ViewCommands.cpp`.
- `source-snapshots\Specs_Plans_WIP_Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md`.
- `source-snapshots\.agents_skills_red-salamander-crash-forensics_SKILL.md`.
- Crash dump and `.txt` sidecar artifacts copied from `C:\Users\eric\AppData\Local\RedSalamander\Crashes`.

## Resume Plan

1. Continue with systematic debugging. Start from the speed-limit prompt crash, not the earlier NavigationView DPI failure.
2. Inspect these implementations before patching:
   - `FileOperationsSpeedLimitPromptWindow::WndProc`, OK, Cancel, close, and `WM_NCDESTROY` paths.
   - `DebugCancelFileOperationsSpeedLimitPrompt`.
   - `DebugSetFileOperationsSpeedLimitPromptText`.
   - `DebugGetFileOperationsSpeedLimitPromptSnapshot`.
   - `GetFileOperationsSpeedLimitPromptHandle`.
   - `InvokeVisibleDescendantByName` / value-setting helpers used by `TestFileOperationsSpeedLimitPromptLiveDxInteraction`.
3. Reproduce narrowly before changing code:
   - `cmd_pane_fileops_speedLimit_prompt_live_dx_interaction`
   - If focused passes, run `cmd_pane_fileops_speedLimit_prompt_uses_dxui_surface,cmd_pane_fileops_speedLimit_prompt_live_dx_interaction`.
   - If that passes, widen from the archived last-run trace toward the FileOps issues-pane predecessor block.
4. If evidence confirms a destructive cross-thread prompt close/confirm path, patch narrowly so teardown-producing actions are posted to the prompt/UI thread rather than synchronously invoked from the worker thread.
5. Rebuild `RedSalamander` Debug, then rerun focused, predecessor/window, and broad Commands.
6. After broad Commands is green, recheck `origin/master...HEAD` (`25` ahead / `4` behind at this archive), preserve/stage current work, resync/rebase, and rerun the post-rebase gates.
7. Final closeout still needs proper Full, final Debug plus test-enabled Release FolderView perf matrices from current HEAD, durable spec migration, and then moving the WIP plan to Done.
