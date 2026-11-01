# FolderView Warp Drive Continuation - 2026-07-01 22:47

## Current State

- Workspace: `Z:\src\RedSalamander`
- Branch: `codex/folderview-warpdrive`
- Origin check: `origin/master...HEAD` = `0 25` at this pause; no rebase was needed.
- Latest clean build: `msbuild-20260701_222239_273.log`, Debug `RedSalamander`, 0 warnings / 0 errors.
- Do not drop the existing `stash@{0}: autostash` noted in earlier archives.

## What Changed In This Session

- Hardened `cmd_pane_navigationView_edit_mode_survives_external_refresh` so it waits for the NavigationView path model to settle, enters edit mode through the real `CommandFocusAddressBar` path first, and only falls back to the synthetic region-key route if needed.
- Hardened the FolderView device-loss and DPI repaint assertions so they use the monotonic `fullClientRenderCount` event signal instead of requiring the volatile `lastRenderWasFullClient` snapshot to still be the most recent render in broad-suite conditions.
- No product logic was changed in this session; only `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp` was edited after the previous staged state.

## Verification Captured Here

- Build passed:
  - `msbuild-20260701_222239_273.log`
- Focused green:
  - `folderView_render_device_loss_recovers`: `focused-device-loss-run-all-tests-results.json`
  - `folderView_dpi_change_repaints_both_panes`: `focused-dpi-repaint-run-all-tests-results.json`
  - `cmd_pane_navigationView_edit_mode_survives_external_refresh`: `focused-nav-edit-refresh-run-all-tests-results.json`
  - `cmd_pane_filter_prompt_live_dx_interaction`: `focused-filter-prompt-live-dx-run-all-tests-results.json`
- Broad Commands still red:
  - `commands-broad-after-repaint-assertion-run-all-tests-results.json`
  - Wrapper exit: `1`
  - App exit: `-1073741819` (`0xC0000005`)
  - WER: `textinputframework.dll`, exception `0xc0000005`, fault offset `0x000000000006a419`
  - Trace stops in `cmd_pane_filter_prompt_live_dx_interaction` after `pane_filter_prompt_live_dx: cycle 'live Cancel interaction' worker-done`, before the case returns.
- Reduced prompt-heavy slice passed:
  - Case slice: indexes 602..627, saved as `prompt-slice-602-627.filter.txt`
  - Result: `26 passed / 0 failed / 0 skipped`, `commands-prompt-slice-602-627-run-all-tests-results.json`

## Interpretation

- Origin does not contain a fix to rebase onto.
- The previous navigation edit-mode test issue is fixed and focused-green.
- The previous broad-only FolderView repaint assertion issue is fixed/hardened and focused-green.
- The remaining blocker is the suite-order/timing-dependent native text input / TSF crash in `textinputframework.dll`.
- The crash is not reproduced by the focused filter prompt case or by the local 602..627 prompt slice, so continue with reduced prefix/suite-order reproduction rather than patching the focused case.

## Resume Plan

1. Keep the staged TSF/native text work intact; do not revert unrelated staged changes.
2. Inspect `Common/DxUi/DxUi.TextStoreACP.cpp` around sink/advice, `RequestLock`, `DisconnectNativeTextInputTextStore`, and `DetachNativeTextInputTextStore`.
3. Inspect modal prompt close paths in `RedSalamander/FolderWindow.FileSystem.cpp`, especially filter/change-case/create-directory prompts and their `_dxHost.SetFocusControl(nullptr)` before `_hWnd.reset()` guards.
4. Build a tighter non-traced failing replay:
   - Start from broader slices that include earlier FileOps/issues-pane and navigation-shell cases, not just the local 602..627 prompt slice.
   - Use comma-separated exact `--selftest-case=` filters generated from `--selftest-list-cases`.
5. If a reduced slice crashes, rerun the same slice once with `REDSALAMANDER_DXUI_TEXTINPUT_TRACE_FILE` pointed at a file under `.build/logs/`, but treat a traced pass as timing-perturbed evidence.
6. Do not move the WIP plan to Done until broad Commands, FileOps, ViewerPE/ViewerVLC HUD, Tools Pester, Full skip-build, final Debug plus test-enabled Release FolderView perf matrices, and spec migration are all green and cited.

## Important Files In This Archive

- `staged-diff.patch`
- `unstaged-viewcommands-diff.patch`
- `git-status-short-final.txt`
- `git-diff-stat-final.txt`
- `git-diff-cached-stat-final.txt`
- `commands-broad-crash-commands-trace.txt`
- `commands-broad-crash-last-run-trace.txt`
- `application-error-redsalamander-recent.txt`
- `windows-error-reporting-recent.txt`
- `prompt-slice-602-627.filter.txt`
- `untracked-wip-files-snapshotted.txt`
- `untracked-wip-snapshots/`
