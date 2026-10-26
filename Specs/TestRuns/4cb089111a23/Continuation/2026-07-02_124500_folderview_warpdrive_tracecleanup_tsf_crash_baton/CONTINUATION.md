# FolderView WarpDrive Continuation Baton - 2026-07-02 12:45

## Current State

- Branch: `codex/folderview-warpdrive`
- HEAD at archive time: `275c040348a3994d45c6aef340333a35a0490cb6`
- Origin delta at archive time: `25 0` from `origin/master...HEAD`
- The worktree is intentionally dirty and heavily staged from the ongoing WarpDrive closeout.

## What Changed Since The 11:48 Baton

- The focused Edit New validation failure was a selftest race, not a product validation miss. `DebugConfirmFolderViewEditNewPrompt()` posts the confirm command, so the worker could snapshot before the UI thread applied validation. The test now polls until the prompt closes or a non-empty validation message appears, and failure output names the exact missing inputs.
- Temporary Archive Pack prompt trace diagnostics were removed from `RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp` and `RedSalamander/SelfTest/Commands/Commands.SelfTest.Dialogs.cpp`.
- Source search confirms no remaining `archive_pack_prompt:`, `archive_pack_modal_cycle:`, `archive_pack_test:`, `TraceArchivePackPrompt`, `ArchivePackPromptDebugCommandName`, or old icon hot-path diagnostic strings in `RedSalamander`, `Common`, or `Tests`.

## Green Evidence

- Debug build after the Edit New fix: `run-artifacts/msbuild-20260702_115720_592.log`.
- Focused Edit New after the async-confirm wait: `run-artifacts/editnew_invalid_focused_after_async_confirm_wait_20260702.log`.
- Edit New pair after the async-confirm wait: `run-artifacts/editnew_pair_after_async_confirm_wait_20260702.log`.
- Broad Commands before Archive Pack trace cleanup, still with temporary Pack tracing present: `run-artifacts/commands_broad_after_editnew_async_confirm_wait_20260702.log` passed `776 passed / 0 failed / 2 skipped`.
- Debug build after Archive Pack trace cleanup: `run-artifacts/msbuild-20260702_121851_319.log`.
- Focused Pack prompt after trace cleanup: `run-artifacts/pack_prompt_focused_after_trace_cleanup_20260702.log`.

## Current Blocker

The no-trace broad Commands rerun crashed:

- Wrapper log: `run-artifacts/commands_broad_after_trace_cleanup_editnew_wait_20260702.log`
- Wrapper exit file: `run-artifacts/commands_broad_after_trace_cleanup_editnew_wait_20260702.exit.txt` (`1`)
- Selftest child exit in `last-run/run-all-tests-results.json`: `-1073741819` (`0xC0000005`)
- WER faulting module: `textinputframework.dll`
- Fault offset: `0x000000000006a419`
- Latest dump pointer: `C:\Users\eric\AppData\Local\CrashDumps\RedSalamander.exe.73980.dmp`
- Last trace stops inside `cmd_pane_pack_prompt_uses_dxui_unique_archive_and_packer_extensions`, after `cmd_pane_changeAttributes_options_dialog_uses_dxui_not_win32_template` and `cmd_pane_makeFileList_options_dialog_uses_dxui_not_win32_template`.
- `commands/results.json` was not written because the child crashed before normal selftest closeout.

Do not reintroduce the removed trace spam as the fix. The earlier broad pass with traces is useful evidence that timing/order changed, not that trace writes are a durable guard.

## Resume Plan

1. Inspect the no-trace Pack prompt debug-close lifecycle first. Compare `DebugConfirmArchivePackPrompt()` / `DebugCancelArchivePackPrompt()` with the already-hardened posted close pattern used by FolderView prompt close debug commands.
2. Reproduce the smallest no-trace order window before another broad run:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_pane_editNew_prompt_creates_file_without_applicable_editor,cmd_pane_changeAttributes_options_dialog_uses_dxui_not_win32_template,cmd_pane_makeFileList_options_dialog_uses_dxui_not_win32_template,cmd_pane_pack_prompt_uses_dxui_unique_archive_and_packer_extensions
```

3. If the four-case window is green, widen from `last-run/commands/trace.txt` around the Pack prompt cluster rather than rerunning the full suite blind.
4. If a TSF crash reproduces, preserve `SelfTest\last_run`, WER, and the crash-dump pointer before running any other selftest.
5. After the Pack prompt no-trace blocker is fixed, rerun:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2
```

Then continue the existing closeout path: Full suite, final Debug and test-enabled Release FolderView perf matrices, durable spec migration, and move the WIP plan to `Specs/Plans/Done/` only after all gates are green and cited.

