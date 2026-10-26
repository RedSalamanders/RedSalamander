# FolderView WarpDrive Continuation Baton

Archive: Specs/TestRuns/4cb089111a23/Continuation/2026-07-03_114630_folderview_warpdrive_preferences_keyboard_baton
Created: 2026-07-03 11:47:07 +02:00
Branch/HEAD:
codex/folderview-warpdrive
275c040348a3994d45c6aef340333a35a0490cb6

## Current State

- The plan remains WIP: Specs/Plans/WIP/Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md.
- Do not move it to Done yet. Task 10 Step 4 and Task 11 remain open, and the closeout checklist still needs final full-suite plus final Debug/test-enabled Release FolderView perf evidence cited from the current HEAD.
- Latest clean build after the current nav/menu hardening: .build/codex-runs/msbuild_redsalamander_after_nav_menu_waits_20260703.log (successful Debug build, 0 warnings / 0 errors per prior run output).

## What Changed Since The Previous Baton

- Hardened cmd_pane_navigationView_path_doubleClick_after_focusAddressBar_tab_traversal by waiting for a real edit-surface focus target before sending Tab.
- Hardened cmd_app_menuBar_persistent_view_to_plugins_hover_switches_popup by waiting for the replacement popup and Plugins selected index to settle, instead of stopping at the first root-switch observation.
- Rebuilt Debug and verified the two focused cases plus the cmd_pane_navigationView_ and cmd_app_menuBar_ prefixes.

## Latest Evidence

- Focused nav/menu summary: run-artifacts/runall_commands_focused_after_nav_menu_waits_summary_20260703.json.
- NavigationView prefix: run-artifacts/runall_commands_navigationView_prefix_after_nav_menu_waits_20260703.run-all.json (16 passed / 0 failed).
- MenuBar prefix: run-artifacts/runall_commands_menuBar_prefix_after_nav_menu_waits_20260703.run-all.json (17 passed / 0 failed).
- Broad Commands after nav/menu patch, run 1: run-artifacts/commands_broad_after_nav_menu_waits_20260703.run-all.json (775 passed / 1 failed / 2 skipped); failure was cmd_pane_navigationView_path_region_keyboard_activation_enters_edit_mode.
- Focused nav keyboard and exact predecessor pair both passed afterward: run-artifacts/runall_commands_nav_path_region_keyboard_focused_after_broad_red_20260703.run-all.json and run-artifacts/runall_commands_nav_doubleclick_then_keyboard_pair_after_broad_red_20260703.run-all.json.
- Broad Commands after nav/menu patch, run 2: run-artifacts/commands_broad_after_nav_menu_waits_rerun2_20260703.run-all.json (775 passed / 1 failed / 2 skipped). This time NavigationView passed; the single current red is cmd_preferences_dialog_keyboard_long_run_list_scrolling_stays_bounded.

See failure-summary.txt for machine-readable JSON summaries and failure reasons.

## Current Red Point

Current failure from broad Commands rerun 2:

cmd_preferences_dialog_keyboard_long_run_list_scrolling_stays_bounded

Reason:

Preferences category tree host should not churn repainting during Keyboard list scrolling chunk 4; render count moved from 2 to 4.

The related source is snapshotted here:

- source-snapshots/RedSalamander_SelfTest_Commands_Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp
- source-snapshots/RedSalamander_SelfTest_Commands_Commands.SelfTest.Preferences.cpp

## Recommended Resume Steps

1. Read the keyboard long-run test around Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp lines ~1016-1094 and helper WaitForPreferencesCategoryTreeRenderCountToSettle in Commands.SelfTest.Preferences.cpp.
2. Run the focused case: .\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_preferences_dialog_keyboard_long_run_list_scrolling_stays_bounded -TimeoutMultiplier 2.
3. Identify the immediate predecessor window from run-artifacts/commands_broad_after_nav_menu_waits_rerun2_20260703.run-all.json, then run the predecessor cluster plus target if the focused case passes.
4. Decide whether the assertion is too strict after broad-suite focus/hover invalidation, or whether product category-tree repaint churn is real. Patch only after that distinction is proven.
5. Rebuild Debug if patched, rerun the focused Preferences case, the cmd_preferences_dialog_ prefix if appropriate, then full broad Commands.
6. Once broad Commands is green, continue with ViewerPE direct, Monitor ETW direct, .\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild, and final Debug plus test-enabled Release FolderView perf matrices.
7. Update the WIP spec evidence log after stable green gates. Move to Done only when Task 11 closeout is truly complete.

## Archive Contents

- git-status-short.txt, branch/head, origin ahead/behind, staged and unstaged diff stats/name-status excluding Specs/TestRuns/**.
- patches/staged-code-and-spec-no-testruns.patch and patches/unstaged-code-and-spec-no-testruns.patch.
- Relevant .build/codex-runs logs/results/traces for the current nav/menu and broad Commands runs.
- WIP spec snapshot plus current source snapshots for nav/menu hardening and the next Preferences investigation.
- Latest RedSalamander crash sidecar .txt files from %LOCALAPPDATA%\RedSalamander\Crashes.