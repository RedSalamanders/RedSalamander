# FolderView WarpDrive Continuation - Edit New Baton

Created: 2026-07-02 11:48:47 local time
Branch: `codex/folderview-warpdrive`
HEAD: `275c040348a3994d45c6aef340333a35a0490cb6`
Origin/master: `45ae2a9a90211aa71a961178661db3dfc36f9876`
Ahead/behind origin/master: `25 0`

## Current State

The worktree is still WIP. Do not move the plan to Done yet.

The latest broad Commands investigations narrowed past several transient or order-sensitive failures:

- `cmd_compare_directories_options_tab_traversal_live_dx_interaction` passed focused, in its immediate predecessor cluster, and in the archived case window `506..522`.
- `cmd_plugin_configuration_dialog_tab_traversal_live_dx_interaction` passed focused and in its 7-case predecessor cluster.
- `cmd_connection_credential_prompt_live_dx_interaction` passed focused, in its 5-case credential cluster, in the exact `178..213` window, and in the full `0..213` prefix.

The current actionable blocker is stable and local:

- Broad Commands rerun `commands_broad_after_prefix213_green_20260702.log` failed after `641 passed / 1 failed / 136 skipped`.
- Failing case: `cmd_pane_editNew_prompt_rejects_invalid_and_existing_names`.
- Failure reason: `Edit New should show a validation message for every invalid file name.`
- Focused rerun `editnew_invalid_focused_after_broad_red_20260702.log` fails the same way in 391 ms.

No `RedSalamander.exe` or `RedSalamanderMonitor.exe` process was left running when this archive was created.

## Evidence Included

Core state:

- `git-status-short.txt`
- `git-diff-cached-name-status.txt`
- `git-diff-name-status.txt`
- `git-diff-cached-stat.txt`
- `git-diff-stat.txt`
- `patches/staged-code-and-spec-no-testruns.patch`
- `patches/unstaged-diff-no-testruns.patch`
- `process-snapshot-redsalamander-dxui.txt`
- `Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.snapshot.pre-status.md`

Run artifacts:

- `run-artifacts/commands_broad_after_prefix213_green_20260702.log`
- `run-artifacts/commands_broad_after_prefix213_green_20260702/run-all-tests-results.json`
- `run-artifacts/commands_broad_after_prefix213_green_20260702/commands/results.json`
- `run-artifacts/commands_broad_after_prefix213_green_20260702/commands/trace.txt`
- `run-artifacts/editnew_invalid_focused_after_broad_red_20260702.log`
- `run-artifacts/editnew_invalid_focused_after_broad_red_20260702/run-all-tests-results.json`
- `run-artifacts/editnew_invalid_focused_after_broad_red_20260702/commands/results.json`
- `run-artifacts/editnew_invalid_focused_after_broad_red_20260702/commands/trace.txt`
- `run-artifacts/commands_broad_after_prefix213_green_editnew_excerpt.txt`
- `run-artifacts/editnew_invalid_focused_excerpt.txt`

Classification logs kept for context:

- `run-artifacts/compare_options_tab_focused_after_prefs_root_state_fix_20260702.log`
- `run-artifacts/compare_options_tab_predecessor6_after_prefs_root_state_fix_20260702.log`
- `run-artifacts/compare_options_tab_window506_522_after_prefs_root_state_fix_20260702.log`
- `run-artifacts/plugin_config_tab_focused_after_compare_window_green_20260702.log`
- `run-artifacts/plugin_config_tab_predecessor7_after_compare_window_green_20260702.log`
- `run-artifacts/credential_prompt_live_focused_after_plugin_config_green_20260702.log`
- `run-artifacts/credential_prompt_live_predecessor5_after_plugin_config_green_20260702.log`
- `run-artifacts/credential_prompt_live_window178_213_after_broad_red_20260702.log`
- `run-artifacts/credential_prompt_live_prefix0_213_after_broad_red_20260702.log`

Source snapshots:

- `source-snapshots/Commands.SelfTest.Dialogs.cpp`
- `source-snapshots/FolderWindow.FileSystem.Commands.Part.cpp`
- `source-snapshots/Preferences.Dialog.cpp`
- `source-snapshots/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp`
- `source-snapshots/Commands.SelfTest.CompareOptions.cpp`
- `source-snapshots/Commands.SelfTest.PluginConfig.cpp`
- `source-snapshots/Commands.SelfTest.Connections.cpp`

## Repro Commands

Focused current red:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_pane_editNew_prompt_rejects_invalid_and_existing_names
```

Direct executable equivalent from the captured log:

```powershell
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-fail-fast --selftest-timeout-multiplier=2 --selftest-case=cmd_pane_editNew_prompt_rejects_invalid_and_existing_names
```

Broad rerun after the focused fix:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2
```

Do not run multiple `Run-AllTests.ps1` invocations in parallel. They share `SelfTest\last_run`.

## Next Step

Start with the focused red. The focused trace only records the final assertion, so inspect or temporarily improve diagnostics around:

- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Dialogs.cpp:6947`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Dialogs.cpp:6975`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Dialogs.cpp:6982`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Dialogs.cpp:6999`

The test accumulates `missingValidationForNames`, but the assertion does not currently print which names are missing validation. If code inspection does not immediately identify the product bug, add diagnostic output to name the missing invalid/existing-name probes, rebuild Debug, and rerun the focused case before changing product behavior.

Likely files to inspect first:

- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Dialogs.cpp`
- `RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp`

After the focused case is green, run:

1. The focused Edit New case.
2. The predecessor pair including `cmd_pane_editNew_prompt_filters_editor_combo_and_creates_file`.
3. Broad Commands.
4. Then continue the closeout chain: Full suite and final Debug plus test-enabled Release FolderView perf matrices.

## Do Not Forget

- Keep `Specs/Plans/WIP/Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md` updated with each gate.
- Do not move the plan to `Specs/Plans/Done/` until broad Commands, Full, final perf evidence, and spec migration are all green and cited.
- Preserve unrelated unstaged user/worktree changes. At archive time, unstaged edits still existed in:
  - `RedSalamander/FolderWindow.FileOperations.Popup.cpp`
  - `RedSalamander/FolderWindow.FileSystem.Navigation.Part.cpp`
  - `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`
  - `Tests/DxUiTests/DxUiTests.Controls.cpp`
  - `Tests/DxUiTests/DxUiTests.NativeTextInput.cpp`
