# FolderView WarpDrive Continuation Baton - 2026-07-03 21:49

This archive supersedes:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-03_205410_folderview_warpdrive_compare_green_viewerpe_hud_baton\
```

Current branch/head:

```text
branch: codex/folderview-warpdrive
head:   275c040348a3994d45c6aef340333a35a0490cb6
```

Use `plan-code-test-perf-loop` for the WIP-plan loop and `systematic-debugging`
for the remaining failures. Do not apply another product/test fix until the
specific failing boundary has been reproduced or otherwise explained.

## What Changed Since The 20:54 Archive

- ViewerPE VLC HUD diagnosis was added in
  `Tests\ViewerPETests\ViewerPETests.cpp`:
  - added `PrintChildWindowDiagnostics(HWND, std::wstring_view)`;
  - added a focused debug snapshot dump when the VLC HUD window is missing;
  - added `<cstdint>` for the diagnostic helper.
- `ViewerPETests` Debug rebuild passed:
  `run-artifacts\msbuild_viewerpetests_after_vlc_hud_diag_20260703.log`.
- Focused ViewerPE VLC reruns passed:
  - `run-artifacts\viewerpe_vlc_tab_focus_focused_with_child_diag_20260703.log`
  - `run-artifacts\viewerpe_vlc_tab_focus_focused_with_child_diag_rerun_20260703.log`
- Focused shell combo passed:
  `run-artifacts\viewerpe_shell_combo_longrun_focused_after_vlc_rebuild_20260703.log`.
- Direct broad ViewerPE passed before the proper Full:
  `run-artifacts\viewerpe_direct_broad_after_focused_green_20260703.log`.
- Pester inventory drift was corrected:
  - `Tools\Tests\TestInventory.Tests.ps1`: CompareDirectories registration count
    `229` -> `230`;
  - `Specs\Testing\Testing_TestCoverage.md`: CompareDirectories source fallback
    count `229` -> `230`.
- Focused TestInventory Pester passed:
  `run-artifacts\pester_testinventory_after_compare_count_20260703.log`.
- Tools Pester passed with the explicit `$result.FailedCount` wrapper:
  `run-artifacts\pester_tools_full_after_compare_count_exitwrapped_summary_20260703.log`
  reports `Passed=115 Failed=0 Skipped=0 Total=115`.
  The earlier `pester_tools_full_after_compare_count_20260703.log` had no Pester
  failures but returned process exit `7` because a child negative-test process
  polluted `$LASTEXITCODE`; do not use that process code as red evidence.

## Current Verification State

The proper Full run is archived here:

```text
run-artifacts\full_after_compare_viewerpe_pester_green_20260703.log
last-run-top-level\run-all-tests-results.json
```

It ran:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2 *> .build\codex-runs\full_after_compare_viewerpe_pester_green_20260703.log
```

Result:

```text
exit:    1
passed:  1098
failed:  2
skipped: 52
total:   1152
```

Green in that Full run:

- CompareDirectories: 208 passed, 0 failed, 30 skipped.
- Commands: 776 passed, 0 failed, 2 skipped.
- FileOperations: 102 passed, 0 failed, 20 skipped.
- ToolsPester: 115 passed, 0 failed, 0 skipped.

Still red:

- `DxUiTests`
- `ViewerPETests`

## Red Gate 1: DxUiTests Menu

Latest Full `last-run-top-level\DxUiTests.output.log` ends with:

```text
116:  [START] TestMenuBarHoverMessageSwitchesRootWhilePopupOverlapsMenuBar
117:FAILED: overlapping View popup exposes debug state
```

The failing test is in:

```text
Tests\DxUiTests\DxUiTests.Menu.cpp
```

Immediate files/helpers to inspect:

- `WaitForContextMenuPopupState` around `DxUiTests.Menu.cpp:431`.
- `WaitForOwnedContextMenuPopupWindowByFirstItemText`.
- `DebugGetContextMenuPopupState` declaration in `Common\DxUi\DxUi.h`.
- `DebugGetContextMenuPopupState` implementation in
  `Common\DxUi\DxUi.Menu.cpp`.
- Popup debug message handling/state fill around `DxUi.Menu.cpp` menu popup
  window procedure and debug-state helpers.

Suggested first reproduction:

```powershell
.\.build\x64\Debug\DxUiTests.exe --suite=Menu *> .build\codex-runs\dxui_menu_suite_after_full_red_20260703.log
```

If it does not reproduce, run direct broad:

```powershell
.\.build\x64\Debug\DxUiTests.exe *> .build\codex-runs\dxui_direct_broad_after_full_red_20260703.log
```

Current root-cause lead is timing/registration readiness rather than a proven
product failure: the driver finds the View popup HWND and immediately waits for
debug state, but the debug-state query can return false under the full-suite
schedule. Prove where the false comes from before changing the wait or the
product debug hook.

## Red Gate 2: ViewerPETests Fresh VLC Harness

Latest Full `last-run-top-level\ViewerPETests.output.log` shows:

```text
317:[PASS] TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly harness path resolves
318:[PASS] TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly fresh harness kill-on-close job object is created
319:[PASS] TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly fresh harness kill-on-close job object is configured
320:[PASS] TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly fresh harness process launches
321:[PASS] TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly fresh harness process joins the kill-on-close job
322:[PASS] TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly fresh harness process resumes
323:[PASS] TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly fresh harness process exits before timeout
324:[PASS] TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly fresh harness process reports an exit code
325:[FAIL] TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly fresh harness run passes
```

Important contrast: focused VLC and direct broad ViewerPE passed before the Full
run. Reproduce the current post-Full state before patching:

```powershell
.\.build\x64\Debug\ViewerPETests.exe *> .build\codex-runs\viewerpe_direct_broad_after_full_red_20260703.log
.\.build\x64\Debug\ViewerPETests.exe TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly *> .build\codex-runs\viewerpe_vlc_tab_focus_after_full_red_20260703.log
```

Then inspect:

```text
Tests\ViewerPETests\ViewerPETests.cpp
```

Search for `RunFilteredSelfExecutable` and the
`TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly fresh harness run passes`
assertion. The failure is in the parent fresh-harness wrapper, not necessarily in
the direct focused child. Earlier evidence also showed one focused child pass with
zero bytes of output and exit `0`, so do not assume missing child output means a
failure.

## Current Source/Diff State To Preserve

New or changed files from this checkpoint that matter for resume:

- `Tests\ViewerPETests\ViewerPETests.cpp`
- `Tools\Tests\TestInventory.Tests.ps1`
- `Specs\Testing\Testing_TestCoverage.md`

This archive includes focused staged/unstaged patches:

```text
patches\focused-staged-diff.patch
patches\focused-unstaged-diff.patch
```

It also includes source snapshots for the ViewerPE test, DxUi menu test and menu
implementation, the WIP plan, inventory Pester, test coverage doc, and
`Tests\README.md`.

There are many pre-existing staged and unstaged changes in this worktree. They
are intentional work-in-progress state. Do not reset, unstage, or revert unrelated
files while continuing this task.

## Resume Plan

1. Reproduce and root-cause the DxUi Menu failure first or prove it no longer
   reproduces under `--suite=Menu` plus broad direct DxUi.
2. Reproduce the ViewerPE fresh VLC wrapper failure after the Full run; determine
   whether the direct child is red or only the parent `RunFilteredSelfExecutable`
   assertion is red.
3. Apply one minimal fix at a time, with a focused red/green guard and no broad
   test masking.
4. Rebuild only what changed, then rerun focused cases, direct broad DxUi,
   direct broad ViewerPE, and Tools Pester if inventory docs change again.
5. Rerun:

   ```powershell
   .\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2
   ```

6. Only after Full is green, rerun and archive the final Debug and test-enabled
   Release FolderView perf matrices, update durable specs, and move the WIP plan
   to `Specs\Plans\Done\`.

Do not mark the active goal complete yet. The plan remains blocked on the two
Full-suite red gates plus final perf/spec closeout.
