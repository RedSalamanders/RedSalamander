# FolderView WarpDrive Continuation Baton - 2026-07-03 08:40

## Current State

- Branch: `codex/folderview-warpdrive`
- HEAD: `275c04034`
- Origin sync: `25 0` from `git rev-list --left-right --count HEAD...origin/master`; no rebase is currently needed.
- Active WIP plan: `Specs\Plans\WIP\Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md`
- This is still Task 11 closeout / convergence work. Do not move the WIP plan to Done until Full is green and the final Debug plus test-enabled Release FolderView perf matrices are rerun and cited from the current HEAD.

## What Changed Since The Previous Baton

Previous baton:

- `Specs\TestRuns\4cb089111a23\Continuation\2026-07-03_083500_folderview_warpdrive_compare_green_viewerpe_red_baton\CONTINUATION.md`

New confirmations and fixes in this slice:

- Fixed the real `ViewerPETests.exe` failure by making `TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly` target the actual `RedSalamander.ViewerVLC.Video` child window instead of whichever tab-stop child happened to have focus.
- Added explicit ViewerVLC video/HUD class constants in `Tests\ViewerPETests\ViewerPETests.cpp`.
- Verified the ViewerVLC focus case was flaky before the fix: 1 failure in 10 repeat runs.
- Verified it was stable after the fix: 10/10 repeat runs green.
- Verified full direct `ViewerPETests.exe` is now green.
- Fixed `RedConfigureTests` standalone build by including `<objidl.h>` before `<XmlLite.h>` in `RedConfigure\Workspace\WorkspaceDiscovery.cpp`.
- Verified `RedConfigureTests.exe` now builds and runs green.
- Re-ran the exact Full-plan `RedSalamanderMonitorEtwLatency` command directly; it exits 0 now, with no code change.
- Isolated `ToolsPesterTests` to stale inventory counts, updated the source/docs counts, and verified the focused `TestInventory.Tests.ps1` Pester file is green.

## Verification Performed

ViewerPE focused flaky repro before fix:

```powershell
.\.build\x64\Debug\ViewerPETests.exe TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly *> .build\codex-runs\viewerpe_vlc_focus_repeat10_20260703.log
```

Result: repeat loop showed 1 failure in 10 runs. Failure was the VLC video surface not owning focus before Tab routing.

ViewerPE build after fix:

```powershell
.\build.ps1 -ProjectName ViewerPETests -Configuration Debug *> .build\codex-runs\msbuild_viewerpe_focus_harness_20260703.log
```

Result: exit 0, 0 warnings/errors.

ViewerPE focused repeat after fix:

```powershell
.\.build\x64\Debug\ViewerPETests.exe TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly *> .build\codex-runs\viewerpe_vlc_focus_repeat10_after_video_target_20260703.log
```

Result: 10/10 green, `failures=0`.

ViewerPE full direct after fix:

```powershell
.\.build\x64\Debug\ViewerPETests.exe *> .build\codex-runs\viewerpe_direct_after_vlc_focus_target_20260703.log
```

Result: exit 0, `ViewerPETests passed.`

RedConfigureTests failing build before include-order fix:

```powershell
.\build.ps1 -ProjectName RedConfigureTests -Configuration Debug *> .build\codex-runs\msbuild_redconfiguretests_20260703.log
```

Result: exit 1. `WorkspaceDiscovery.cpp` failed in `<XmlLite.h>` because COM interface declarations were unavailable.

RedConfigureTests build after fix:

```powershell
.\build.ps1 -ProjectName RedConfigureTests -Configuration Debug *> .build\codex-runs\msbuild_redconfiguretests_after_xmllite_order_20260703.log
```

Result: exit 0, 0 warnings/errors.

RedConfigureTests direct after fix:

```powershell
.\.build\x64\Debug\RedConfigureTests.exe *> .build\codex-runs\redconfiguretests_after_xmllite_order_20260703.log
```

Result: exit 0, `RedConfigureTests passed.`

Monitor ETW direct recheck:

```powershell
.\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --wait-instance --perf --monitor-etw-burst-mode=latency --monitor-etw-burst-count=60 --monitor-etw-burst-size=260 *> .build\codex-runs\monitor_etw_latency_direct_20260703.log
```

Result: exit 0, empty stdout/stderr.

Tools Pester direct before inventory-count updates:

```powershell
Invoke-Pester -Path Tools\Tests -PassThru *> .build\codex-runs\tools_pester_direct_20260703.log
```

Result: exit 7, `Passed: 111 Failed: 4 Skipped: 0`. All failures were stale inventory expectations.

Current authoritative inventory:

```powershell
.\Tools\Get-TestInventory.ps1 *> .build\codex-runs\test_inventory_current_20260703.log
```

Current source-derived counts:

- Commands: 676
- CompareDirectories: 229
- FileOperations: 120
- PerformanceTests2: 12
- DxUiTests NativeTextInput: 117
- ToolsPester: 115
- RequiresBuildToolchain: 1
- VcpkgMergeSynthetic: 5
- VcpkgMergeLockValidation: 3

Focused inventory Pester after count updates:

```powershell
Invoke-Pester -Path Tools\Tests\TestInventory.Tests.ps1 -PassThru *> .build\codex-runs\test_inventory_pester_after_counts_20260703.log
```

Result: exit 0, `Passed: 5 Failed: 0 Skipped: 0`.

## Files Edited In This Slice

Unstaged focused files:

- `Tests\ViewerPETests\ViewerPETests.cpp`
- `RedConfigure\Workspace\WorkspaceDiscovery.cpp`
- `RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp`
- `Tools\Tests\TestInventory.Tests.ps1`
- `Specs\Testing\Testing_TestCoverage.md`
- `Tests\README.md`

Important index state:

- `Tools\Tests\TestInventory.Tests.ps1`, `Specs\Testing\Testing_TestCoverage.md`, and `Tests\README.md` already had staged changes before this slice, and now also have unstaged count updates.
- Avoid broad staging. Stage intentionally when ready.

## Archived Artifacts

State and diffs:

- `git-status-short.txt`
- `git-branch.txt`
- `git-head.txt`
- `origin-master-ahead-behind.txt`
- `git-diff-name-status-no-testruns.txt`
- `git-diff-stat-no-testruns.txt`
- `git-diff-cached-name-status-no-testruns.txt`
- `git-diff-cached-stat-no-testruns.txt`
- `patches\unstaged-code-and-spec-no-testruns.patch`
- `patches\staged-code-and-spec-no-testruns.patch`
- `patches\focused-current-unstaged.patch`
- `patches\focused-current-staged.patch`

Run logs:

- `run-artifacts\viewerpe_direct_repro_20260703.log`
- `run-artifacts\viewerpe_vlc_focus_focused_20260703.log`
- `run-artifacts\viewerpe_vlc_focus_repeat10_20260703.log`
- `run-artifacts\msbuild_viewerpe_focus_harness_20260703.log`
- `run-artifacts\viewerpe_vlc_focus_repeat10_after_video_target_20260703.log`
- `run-artifacts\viewerpe_direct_after_vlc_focus_target_20260703.log`
- `run-artifacts\msbuild_redconfiguretests_20260703.log`
- `run-artifacts\msbuild_redconfiguretests_after_xmllite_order_20260703.log`
- `run-artifacts\redconfiguretests_after_xmllite_order_20260703.log`
- `run-artifacts\monitor_etw_latency_direct_20260703.log`
- `run-artifacts\tools_pester_direct_20260703.log`
- `run-artifacts\test_inventory_current_20260703.log`
- `run-artifacts\test_inventory_pester_after_counts_20260703.log`

Snapshots:

- `source-snapshots\Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.snapshot.md`
- `source-snapshots\ViewerPETests.cpp`
- `source-snapshots\WorkspaceDiscovery.cpp`
- `source-snapshots\TestInventory.Tests.ps1`
- `source-snapshots\Testing_TestCoverage.md`
- `source-snapshots\Tests_README.md`
- `source-snapshots\CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp`
- `source-snapshots\red-salamander-crash-forensics.SKILL.md`

Crash/debug context:

- `crash-sidecars\red-salamander-crashes-latest.txt`
- Latest crash sidecar `.txt` files from `C:\Users\eric\AppData\Local\RedSalamander\Crashes`, if present.

## Remaining Gates

Continue in this order:

1. Run full Tools Pester now that the focused inventory file is green:

```powershell
Invoke-Pester -Path Tools\Tests -PassThru *> .build\codex-runs\tools_pester_after_inventory_counts_20260703.log
```

2. Run a fresh Debug build to include the ViewerPE, RedConfigure, Compare, and Pester-count changes:

```powershell
.\build.ps1 -Configuration Debug *> .build\codex-runs\msbuild_debug_after_viewerpe_redconfigure_pester_20260703.log
```

3. Re-run Full from the current build:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2 *> .build\codex-runs\full_skipbuild_after_red_gates_20260703.log
```

4. If Full is green, rerun final FolderView perf matrices in Debug and test-enabled Release from this HEAD.

5. Update `Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md` with final green gates and perf archives.

6. Merge any durable testing/perf contracts into authoritative docs, then move the WIP plan to `Specs\Plans\Done\`.

## Known Status By Gate

- Compare four red gates from the Full baton: focused green, but Full not rerun after all fixes.
- ViewerPE parent failure: fixed; direct full `ViewerPETests.exe` green.
- RedConfigureTests missing exe/build failure: fixed; direct build and test green.
- Monitor ETW latency: direct command green; Full not rerun.
- Tools Pester: root cause fixed for inventory file; focused Pester green; full `Tools\Tests` Pester still needs rerun.
- Final Debug/Release FolderView perf evidence: still pending.

