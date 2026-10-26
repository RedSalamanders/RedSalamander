# FolderView Warp Drive Continuation - Origin/Keyboard Recheck

Timestamp: 2026-07-01 21:30 local.

## Current Git State

- Branch: `codex/folderview-warpdrive`.
- `git fetch origin` completed successfully.
- `HEAD...origin/master` is `25 0`: this branch is 25 commits ahead and 0 behind `origin/master`; no rebase is currently needed.
- `stash@{0}` still exists as a rebase/autostash backup from earlier work; do not drop it unless explicitly requested.
- The only untracked plan files remain the WIP continuation/remediation specs listed in `git-status-short.txt`.

## Fresh Verification

- Clean Debug rebuild passed: `msbuild-20260701_210811_496.log` (0 warnings / 0 errors).
- Focused Keyboard case now passes from the clean binary:
  - `commands-keyboard-live-clean-20260701_211042.out.log`
  - Case: `cmd_preferences_dialog_keyboard_live_search_dx_interaction`
  - Result: 1 passed / 0 failed / 0 skipped.
- Broad Commands fail-fast is still red:
  - `commands-broad-after-keyboard-clean-20260701_211102.out.log`
  - Result: no `commands/results.json`; wrapper reports `RedSalamander.exe` exit `-1073741819` (`0xC0000005`).
  - WER: `textinputframework.dll`, exception `0xc0000005`, fault offset `0x000000000006a419`, report id `6c69fcac-19a3-4df9-b9eb-c79e71a806fd`.
  - Trace stops while entering `cmd_pane_navigation_create_directory_prompt_keeps_navigation_shell_stable`.

## Resume Point

Do not spend more time on the Preferences Keyboard case unless it regresses again; it is green now. Resume with the TSF/native text crash boundary around the create-directory prompt:

1. Reproduce/minimize `cmd_pane_navigation_create_directory_prompt_keeps_navigation_shell_stable` in the broad-order context.
2. Use the existing TSF lifetime/association evidence in `Common/DxUi/DxUi.NativeTextInput.cpp` and `Tests/DxUiTests/DxUiTests.NativeTextInput.cpp`; do not guess another TSF change without a focused failing guard or stronger dump/trace evidence.
3. Remove or formalize the temporary env-gated `REDSALAMANDER_DXUI_TEXTINPUT_TRACE_FILE` instrumentation before final closeout.
4. After broad Commands is green, continue the remaining Done gates: FileOperations, ViewerPE/ViewerVLC HUD, Tools Pester, Full skip-build, final Debug plus test-enabled Release FolderView perf matrices, spec migration, then move the plan to `Specs/Plans/Done/`.

## Archived Evidence

- `application-error-redsalamander.txt`
- `commands-broad-after-keyboard-clean-20260701_211102.out.log`
- `commands-keyboard-live-clean-20260701_211042.out.log`
- `last_run_trace.txt`
- `msbuild-20260701_210811_496.log`
- `run-all-tests-results.json`
- `git-status-short.txt`
- `origin-master-ahead-behind.txt`
- `git-diff-stat.txt`
- `git-diff-cached-stat.txt`
