# FolderView WarpDrive Continuation Baton - TSF Prompt-Close Follow-up

Created: 2026-07-02 08:30 Europe/Paris

## Branch / Sync State

- Worktree: `Z:\src\RedSalamander`
- Branch: `codex/folderview-warpdrive`
- HEAD: `275c040348a3994d45c6aef340333a35a0490cb6`
- `origin/master...HEAD`: `0 25`
- Do not move the WIP plan to Done yet.

## Current Objective

Continue implementing `Specs\Plans\WIP\Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md`.

The immediate blocker is no longer just the original `textinputframework.dll` access violation. A prompt-close mitigation was added and the exact reduced slice stopped crashing, but four ordinary selftest failures remain in that same slice. Treat the work as active WIP.

## Latest Code Delta Since Previous Archive

Unstaged local edits currently exist in:

- `RedSalamander/FolderWindow.FileSystem.Navigation.Part.cpp`
- `Tests/DxUiTests/DxUiTests.NativeTextInput.cpp`

What changed:

- Added `PostFolderViewPromptCloseDebugCommand(...)`.
- Updated the FolderView prompt debug confirm/cancel helpers for pane filter, selection mask, create-directory, edit-new, and change-case prompts to post close commands instead of cross-thread synchronous `SendMessageW`.
- Extended `DxUiTests.NativeTextInput` source guard so those debug close helpers must route through the shared close helper and must not directly call `SendMessageW`.

Rationale:

- The reduced non-traced crash slice implicated worker-thread debug helpers synchronously destroying DxUi/TSF prompt HWNDs while native text input teardown was in flight.
- Posting the close command keeps actual prompt destruction on the modal UI thread's message dispatch path.

## Evidence Captured Here

This archive includes:

- `msbuild-20260702_081436_016.log`
- `run-artifacts/tsf-repro-*`
- `run-artifacts/tsf-trace-*`
- `run-artifacts/postclose-*`
- `Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.snapshot.md`
- `unstaged-prompt-close-and-sourceguard.patch`
- git status/stat/name-status files
- untracked WIP baton snapshots under `wip-snapshots/`

## Verification So Far

### Build

Command:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
```

Result:

- Exit code: `0`
- Log: `.build\logs\msbuild-20260702_081436_016.log`
- MSBuild diagnostics: `0 warnings`, `0 errors`

### Pre-patch reduced repro, no TSF trace

Run directory:

```text
.build\codex-runs\commands_slice_532_627_tsf_repro_20260702_0005
```

Result:

- Exit code: `1`
- App exit: `0xC0000005`
- WER module: `textinputframework.dll`
- Dump noted during investigation: `C:\Users\eric\AppData\Local\CrashDumps\RedSalamander.exe.73392.dmp`

### Pre-patch same slice with TSF trace

Run directory:

```text
.build\codex-runs\commands_slice_532_627_tsf_trace_20260702_0015
```

Result:

- Exit code: `0`
- `96 passed / 0 failed / 0 skipped`
- The trace perturbed timing and avoided the crash, so use the non-traced repro as the authority.

### Post-patch same slice, no TSF trace

Run directory:

```text
.build\codex-runs\commands_slice_532_627_after_postclose_20260702_0817
```

Result:

- Exit code: `1`
- No AV in this run.
- `92 passed / 4 failed / 0 skipped`

Remaining failures:

1. `cmd_pane_fileops_speedLimit_prompt_long_run_open_close_stays_stable`
   - Reason: `ValuePattern` observed `64,0 KBn` instead of expected prefix `64,0 KB` during cycle 0.
2. `cmd_pane_navigation_switch_pane_focus_keeps_navigation_shell_stable`
   - Reason: focus did not move to expected folder view; `focusedFolderView=0x0`.
3. `cmd_pane_navigation_hot_paths_keeps_navigation_shell_stable`
   - Reason: navigation shell did not restore focus after Hot Paths close; `focusedFolderView=0x0`, `focusedWindow=0x0`.
4. `cmd_app_fatal_error_long_run_open_close_stays_stable`
   - Reason: failed to collect live UI Automation pattern statistics for fatal-error dialog during cycle 3.

## Next Steps

1. Inspect the post-close failure context:

```powershell
Get-Content .build\codex-runs\commands_slice_532_627_after_postclose_20260702_0817\commands-trace.txt
Get-Content .build\codex-runs\commands_slice_532_627_after_postclose_20260702_0817\commands-results.json
Get-Content .build\codex-runs\commands_slice_532_627_after_postclose_20260702_0817\run.out.log -Tail 200
```

2. Run the four failing cases individually:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_pane_fileops_speedLimit_prompt_long_run_open_close_stays_stable -TimeoutMultiplier 2.0
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_pane_navigation_switch_pane_focus_keeps_navigation_shell_stable -TimeoutMultiplier 2.0
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_pane_navigation_hot_paths_keeps_navigation_shell_stable -TimeoutMultiplier 2.0
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_app_fatal_error_long_run_open_close_stays_stable -TimeoutMultiplier 2.0
```

3. Classify failures:

- If they pass individually, rerun the `532..627` slice and investigate cross-case state leakage.
- If they fail individually, fix the local root cause before broad testing.
- In particular, verify whether async posted close requires explicit waits in selftest harness paths that previously assumed synchronous close completion.

4. Run the DxUi source guard after any adjustment:

```powershell
.\build.ps1 -ProjectName DxUiTests -Configuration Debug
.\.build\x64\Debug\DxUiTests.exe
```

5. Only after focused failures pass, rerun:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter <same 532..627 case filter> -TimeoutMultiplier 2.0
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2.0
```

6. Continue the plan gates:

- Release build and Release selftests/perf.
- FolderView perf sample-count gates.
- Archive final perf and test evidence.
- Update authoritative specs.
- Move WIP plan to `Specs\Plans\Done\` only when all required gates are satisfied.

## Open Caution

The post-close helper may need refinement. A thread-aware helper that sends only when already on the prompt owner thread and posts cross-thread could preserve existing synchronous behavior where safe. Do not make that change speculatively; first inspect whether the remaining failures are caused by async close timing or by unrelated focus/UIA state.
