# FolderView WarpDrive Continuation Baton - 2026-07-03 18:14

## Current State

- Branch: `codex/folderview-warpdrive`
- HEAD: `275c040348a3994d45c6aef340333a35a0490cb6`
- Relation to `origin/master`: `4` behind / `25` ahead at archive time.
- No live `RedSalamander` / `DxUiTests` process was observed before this archive pass.
- Latest useful archive root:
  `Specs\TestRuns\4cb089111a23\Continuation\2026-07-03_181403_folderview_warpdrive_plugin_config_baton\`

## What Changed Since The 16:29 Baton

- Hardened the issues-pane focus setup and assertion path:
  - `RedSalamander\FolderWindow.FileOperations.IssuesPane.cpp`
  - `RedSalamander\SelfTest\Commands\Commands.SelfTest.FileOps.cpp`
- The selftest now waits for the selected issue, logical grid focus, and native pane HWND focus before closing the issues pane and asserting folder-focus restoration.
- This is test-support hardening around a real focus contract; it does not weaken the product focus assertion.

## Green Evidence

- Debug build passed:
  `.build\codex-runs\msbuild_redsalamander_after_issues_focus_wait_20260703.log`
- Focused issues-pane hide-focus guard passed:
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_164118`
- Issues-pane tail `531..546` passed:
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_164146`
- Compare Options + issues-pane + speed-limit tail `517..552` passed:
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_164607` (36 passed / 0 failed)
- After the 17:10 broad timeout, the connection-manager modeless connect path passed focused/window/prefix:
  - `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_174151`
  - `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_174227`
  - `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_174418`
- The 18:01 broad-run four failures all passed focused/window reruns:
  - Focused red-4 rerun: `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_180208`
  - Preferences window `356..370`: `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_180249`
  - Find window `432..444`: `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_180331`
- Plugin-configuration tab traversal itself passed in isolation and in its local window:
  - Focused: `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_181126`
  - Window `165..175`: `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_181205`

## Red Evidence

- First broad Commands after the issues-pane patch:
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_170344`
  - 772 passed / 4 failed / 2 skipped.
  - The four failures did not reproduce focused or in local windows.
- Broad rerun after those focused/window greens timed out/hung at the connection-manager modeless connect area.
  - Runner log: `.build\codex-runs\commands_broad_rerun_after_red4_windows_green_20260703.log`
  - Manual hang dump pointer: `C:\Users\eric\AppData\Local\RedSalamander\Crashes\RedSalamander-hang-20260703-174154-p62052.dmp`
  - The dump is intentionally not copied into the repo because it is about 1 GB; `crash-artifacts\hang-dump-pointer.txt` preserves path, size, and timestamp.
  - Focused/window/prefix reruns for that area are green, so treat the timeout as order-sensitive unless it reproduces.
- Second broad Commands:
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_180131`
  - 772 passed / 4 failed / 2 skipped.
  - Failures were Preferences Advanced/File Operations roundtrip and two Find grid-layout baseline exposures.
  - These passed focused/window reruns.
- Current best minimized red point:
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_181043`
  - Prefix `0..370` failed after 370 passed / 1 failed.
  - Failure: `cmd_plugin_configuration_dialog_tab_traversal_live_dx_interaction`.
  - Trace detail:
    - Previous case `cmd_plugin_configuration_dialog_live_dx_interaction` passed.
    - Forward tab traversal advanced through indices `0 -> 7`.
    - At the failed step the test expected `Toggle` label `Use virtual-hosted style addressing`.
    - Actual focus snapshot was already `Edit` label `Max keys per request`, with `scrollY=243`.
    - Trace line: `plugin-config tab step miss: step='use virtual-hosted style addressing toggle' expectedKind=4 expectedLabel='Use virtual-hosted style addressing' labelMatches=0 actualFocusKind=2 actualFocusLabel='Max keys per request' ...`

## Current Hypothesis

The remaining blocker is not the original FolderView perf work. It is broad-order DxUi/plugin-configuration focus reliability.

The plugin-config failure looks like a traversal/snapshot race or stale expected-step alignment after the preceding live interaction case. The command-layer focus advance reports success and a non-null `finalFocus`, but the wrapper trace reports `finalFocus=0x0`, and the assertion samples the next edit (`Max keys per request`) when it expected the virtual-hosted toggle.

Do not weaken the assertion. First prove whether the product focus order is wrong after scroll, the selftest samples after an extra focus advance, or the debug wrapper loses the native/logical focused HWND before the wait predicate observes it.

## Next Implementation Move

Start here:

```text
RedSalamander\SelfTest\Commands\Commands.SelfTest.PluginConfig.cpp
Plugins\FileSystemS3\FileSystemS3.h
```

Use these anchors:

```powershell
rg -n "plugin-config advance|tab step miss|sendTab|Use virtual-hosted|Max keys per request" RedSalamander\SelfTest\Commands\Commands.SelfTest.PluginConfig.cpp
rg -n "Use virtual-hosted|Max keys per request" Plugins\FileSystemS3\FileSystemS3.h
```

Patch only after confirming the real failure mode. Likely useful diagnostics:

- Log the pre-step focus label/kind, post-advance target label/kind, and the final snapshot label/kind for each `sendTab`.
- Log the debug focus-order list after the S3 dialog is scrolled to the virtual-hosted region.
- Check whether `DebugAdvancePluginConfigurationDialogFocus(...)` is advancing from the expected host but the wait predicate samples one step later.

Then rerun in this order:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_plugin_configuration_dialog_tab_traversal_live_dx_interaction" -TimeoutMultiplier 2
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter ".build\codex-runs\commands_plugin_config_tab_window_after_prefix_red_20260703.filter.txt" -TimeoutMultiplier 2
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter ".build\codex-runs\commands_prefix_0_370_after_preferences_broad_red_20260703.filter.txt" -TimeoutMultiplier 2
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2
```

Do not run multiple `Run-AllTests.ps1` processes at once; they share `SelfTest\last_run`.

## Closeout Still Required

- Broad Commands must pass cleanly.
- Then rerun the broader outstanding gates from the WIP plan: FileOperations, ViewerPE/ViewerVLC HUD where applicable, Tools Pester, Full skip-build, and final Debug plus test-enabled Release FolderView perf matrices.
- Only move the WIP plan to `Specs\Plans\Done\` after those gates are green or explicitly waived and durable contracts are merged into authoritative specs.

