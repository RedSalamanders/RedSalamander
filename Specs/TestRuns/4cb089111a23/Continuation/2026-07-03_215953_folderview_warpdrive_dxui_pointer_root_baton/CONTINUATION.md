# FolderView WarpDrive Continuation Baton - 2026-07-03 21:59

This archive supersedes:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-03_214957_folderview_warpdrive_full_red_dxui_viewerpe_baton\
```

The 21:49 archive remains the last full red-run archive. This 21:59 baton captures the DxUi root-cause work done after it.

## Current State

- Branch: `codex/folderview-warpdrive`.
- Active WIP spec: `Specs\Plans\WIP\Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md`.
- Do not move the WIP plan to `Done` yet.
- The proper Full gate is still red:
  - `DxUiTests`: `TestMenuBarHoverMessageSwitchesRootWhilePopupOverlapsMenuBar`.
  - `ViewerPETests`: parent fresh harness failure for `TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly fresh harness run passes`.
- Green evidence already collected before this baton:
  - `CompareDirectories`: 208 passed / 0 failed / 30 skipped.
  - `Commands`: 776 passed / 0 failed / 2 skipped.
  - `FileOperations`: 102 passed / 0 failed / 20 skipped.
  - `Tools` Pester: 115 passed / 0 failed.
  - Focused/direct ViewerPE evidence had been green before the Full parent harness red.

## DxUi Menu Root Cause

The DxUi failure is not a stale debug hook problem. The diagnostic probe showed the captured View popup HWND was already destroyed by the time the wait timed out:

```text
FAILED: overlapping View popup exposes debug state (hwnd=0xed0930 isWindow=false visible=false thread=0 process=0 wmNull=false state=false text=false rect=false firstText='' rect=(0.0,0.0,0.0,0.0))
```

The env-gated menu trace showed the product/test interaction:

- The failing test opens the View popup while the real OS cursor is over the owner menu strip, outside the View popup surface.
- The popup owns capture, so Windows delivers a captured `WM_MOUSEMOVE` to the old View popup using the live cursor position.
- That incidental move calls `switchRootFromPointer` and switches the root to Plugins before the test posts its deliberate inside-popup move.
- The original View popup is destroyed, so the test fails when it tries to wait for debug state on the original HWND.

Evidence is in:

```text
run-artifacts\dxui_menu_trace_overlap_red_20260703.log
run-artifacts\dxui_menu_suite_trace_overlap_red_20260703.log
run-artifacts\dxui_menu_suite_with_state_probe_20260703.log
run-artifacts\dxui_menu_suite_after_overlap_drain_probe_20260703.log
```

The narrow drain attempt did not fix it. Root cause is the captured post-open mousemove, not pre-existing queued mouse messages.

## Current Unverified Test Fix

`Tests\DxUiTests\DxUiTests.Menu.cpp` currently has an unverified test-side fix:

- Added `std::atomic_bool pointerRootSwitchArmed{false};`.
- `sessionCallbacks.switchRootFromPointer` returns `std::nullopt` until the driver arms it.
- The test arms pointer-root switching immediately before posting its deliberate inside-popup `WM_MOUSEMOVE`.
- This should ignore incidental live-cursor root switches during setup while still allowing the deliberate probe to fail if the product incorrectly switches root for an inside-popup move.

Temporary diagnostic state still present:

- `DescribeContextMenuPopupDebugProbe(HWND)` helper.
- Verbose `driverFailure = std::format(...)` failure text.

If the fix goes green, remove those temporary diagnostics and restore the concise failure text:

```cpp
driverFailure = "overlapping View popup exposes debug state";
```

Keep `pointerRootSwitchArmed` only if the focused and broad DxUi runs prove it.

## Resume Commands

Start here:

```powershell
.\build.ps1 -ProjectName DxUiTests -Configuration Debug *> .build\codex-runs\msbuild_dxuitests_menu_arm_pointer_probe_20260703.log; $LASTEXITCODE
.\.build\x64\Debug\DxUiTests.exe --suite=Menu *> .build\codex-runs\dxui_menu_suite_after_arm_pointer_probe_20260703.log; $LASTEXITCODE
```

If `--suite=Menu` is green:

```powershell
# Remove temporary DescribeContextMenuPopupDebugProbe and verbose failure text.
.\build.ps1 -ProjectName DxUiTests -Configuration Debug *> .build\codex-runs\msbuild_dxuitests_menu_cleanup_20260703.log; $LASTEXITCODE
.\.build\x64\Debug\DxUiTests.exe --suite=Menu *> .build\codex-runs\dxui_menu_suite_after_cleanup_20260703.log; $LASTEXITCODE
.\.build\x64\Debug\DxUiTests.exe *> .build\codex-runs\dxui_direct_broad_after_menu_fix_20260703.log; $LASTEXITCODE
```

Then return to the ViewerPE Full-only red:

```powershell
.\.build\x64\Debug\ViewerPETests.exe *> .build\codex-runs\viewerpe_direct_broad_after_full_red_20260703.log; $LASTEXITCODE
.\.build\x64\Debug\ViewerPETests.exe TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly *> .build\codex-runs\viewerpe_vlc_tab_focus_after_full_red_20260703.log; $LASTEXITCODE
```

Inspect `Tests\ViewerPETests\ViewerPETests.cpp`, especially `RunFilteredSelfExecutable` around the fresh-child harness path, before changing product code.

## Archive Contents

This baton includes:

- `git-status-short.txt`
- `git-diff-stat.txt`
- `git-diff-cached-stat.txt`
- `git-diff-name-status.txt`
- `git-diff-cached-name-status.txt`
- `patches\unstaged-code-and-spec-no-testruns.patch`
- `patches\staged-code-and-spec-no-testruns.patch`
- `patches\dxui-menu-current-unstaged.patch`
- `source-snapshots\Tests_DxUiTests_DxUiTests.Menu.cpp`
- `source-snapshots\Common_DxUi_DxUi.Menu.cpp`
- `source-snapshots\Tests_ViewerPETests_ViewerPETests.cpp`
- `source-snapshots\Specs_Plans_WIP_Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md`
- copied recent DxUi, ViewerPE, and Full run logs under `run-artifacts\`
- `codex-runs-recent-files.txt`
- `process-snapshot-redsalamander-dxui-viewerpe.txt`

## Final Convergence Gates Still Required

- DxUi focused Menu suite green.
- DxUi direct broad green.
- ViewerPE focused and direct broad green, including fresh child harness behavior.
- Proper `.\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2` green.
- Final Debug FolderView perf matrix archived and cited.
- Test-enabled Release FolderView perf matrix archived and cited.
- Temporary diagnostics removed or formalized.
- Durable spec updates completed.
- WIP plan moved to `Specs\Plans\Done\` only after all gates are cited.

