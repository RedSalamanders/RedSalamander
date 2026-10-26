# FolderView WarpDrive Continuation - 2026-07-03 20:54

## Current Position

- Branch: `codex/folderview-warpdrive`
- HEAD: `275c040348a3994d45c6aef340333a35a0490cb6`
- Relative to `origin/master`: `25` ahead / `4` behind.
- Goal remains active: finish `Specs\Plans\WIP\Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md`, then move it to Done only after Full and the final Debug plus test-enabled Release FolderView perf matrices are green and cited.
- This archive supersedes the 20:23 baton for resume purposes. The 20:23 baton is still useful for the broader Full-red context and the plugin-config scroll-predicate fix.

## What Changed Since The Previous Baton

CompareDirectories was investigated first.

The focused red from the previous Full run reproduced:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter "decision_cache_eviction_budget_pending_wide_tree" -TimeoutMultiplier 2 *> .build\codex-runs\compare_decision_cache_pending_focused_after_full_red_20260703.log
```

Result: failed with:

```text
pending budget: failed to drain pending subtree updates. pendingSubdirUpdates=9 contentPendingCompares=537 contentQueueSize=533 contentInFlightSize=537.
```

Root cause: the test pins the left `keep` subtree but requests a budget/eviction scenario whose right-side containing parent is `spill`. Under the tiny decision-cache budget, the parent needed to consume pending child decisions can be evicted, so pending subdir update draining churns instead of settling. The test intent is to prove pending child decisions are evictable after parent aggregation consumes them, so the containing parent must remain available.

Focused test-only fix in `RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.Cases.RuntimeAndRemote.cpp`:

```cpp
session->SetPinnedFolders(std::filesystem::path(L"keep"), std::filesystem::path(L"spill"));
```

Note: the same file already had prior unstaged work for the quiet pending-drain helper and `worker_shutdown_joins_before_state_teardown`; do not treat the new diff in that file as only this one-line change.

## Fresh Evidence

- Debug `RedSalamander` rebuild passed after the Compare test fix:
  `.build\codex-runs\msbuild_redsalamander_after_compare_pending_pin_spill_20260703.log`
  reports `0 warning(s), 0 error(s)`.
- Focused `decision_cache_eviction_budget_pending_wide_tree` now passes:
  `.build\codex-runs\compare_decision_cache_pending_focused_after_pin_spill_20260703.log`
  reports `1 passed / 0 failed / 0 skipped`.
- Focused `search_service_sqlite_ntfs_traversal_seed_stays_degraded` passes:
  `.build\codex-runs\compare_search_service_ntfs_seed_focused_after_full_red_20260703.log`
  reports `1 passed / 0 failed / 0 skipped`.
- Broad CompareDirectories now passes:
  `.build\codex-runs\compare_broad_after_pending_pin_spill_20260703.log`
  reports `208 passed / 0 failed / 30 skipped`.
- Focused ViewerPE VLC still fails:
  `.build\codex-runs\viewerpe_vlc_tab_focus_focused_after_full_red_20260703.log`
  reports `[FAIL] ViewerVLC exposes a visible HUD child that can take focus`.

## Current Blockers

1. ViewerPE VLC focused failure:

   ```powershell
   .\.build\x64\Debug\ViewerPETests.exe TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly *> .build\codex-runs\viewerpe_vlc_tab_focus_focused_after_full_red_20260703.log
   ```

   Result:

   ```text
   [PASS] ViewerVLC exposes a visible video child that can take initial focus
   [FAIL] ViewerVLC exposes a visible HUD child that can take focus
   [PASS] ViewerVLC exposes visible video and HUD child surfaces
   ```

   Important clue: `Plugins\ViewerVLC\ViewerVLC.h` defines `kHudClassName` as exactly `RedSalamander.ViewerVLC.Hud`, matching `Tests\ViewerPETests\ViewerPETests.cpp`. This is probably not a simple class-name typo. Resume by inspecting `FindFirstChildWindowByClass(...)`, child enumeration depth, visibility/style filtering, parent/owner relationships, timing, and whether the HUD is created but not discoverable from the tested parent at that moment. A useful first patch would add a temporary diagnostic dump of child HWND/class/style/visible/focusable data around the failing check.

2. ViewerPE shell combo Full failure:

   The previous Full run also failed `TestViewerShellComboHostsLongRunOpenCloseStayStable fresh harness run passes`. This focused case was not rerun after the Compare work. Run it next:

   ```powershell
   .\.build\x64\Debug\ViewerPETests.exe TestViewerShellComboHostsLongRunOpenCloseStayStable *> .build\codex-runs\viewerpe_shell_combo_longrun_focused_after_full_red_20260703.log
   ```

3. ToolsPester inventory drift:

   Full log shows:

   ```text
   CompareDirectories static RunCase count drifted. Expected '229' but got '230'.
   Coverage spec CompareDirectories count drifted from source. Expected '230' but got '229'.
   ```

   Likely fix after confirming with focused Pester: update `Tools\Tests\TestInventory.Tests.ps1` CompareDirectories expected count from `229` to `230`, and update the CompareDirectories count in `Specs\Testing\Testing_TestCoverage.md` from `229` to `230`. Check `Tests\README.md` only if its parsed counts include CompareDirectories.

4. Full suite needs a fresh rerun after the Compare fix.

   The last proper Full run before this baton was still red:

   ```powershell
   .\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2 *> .build\codex-runs\full_after_commands_fileops_green_20260703.log
   ```

   Summary from that old run: `320 passed / 4 failed / 50 skipped`. Its two Compare failures should be rechecked by a new Full run because broad Compare is now green.

## Resume Order

1. Add focused diagnostics to `Tests\ViewerPETests\ViewerPETests.cpp` for the VLC HUD failure, or inspect the existing helpers first if enough context is obvious.
2. Rerun `TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly`.
3. Run `TestViewerShellComboHostsLongRunOpenCloseStayStable`.
4. Fix and rerun focused ToolsPester inventory coverage drift.
5. Run direct broad `ViewerPETests.exe` once the two focused cases are green.
6. Run ToolsPester or Full depending on confidence; avoid running two `Run-AllTests.ps1` processes in parallel because they share `SelfTest\last_run`.
7. Run proper Full.
8. Only after Full is green, rerun the final Debug and test-enabled Release FolderView perf matrices from current HEAD and cite the new archives in the WIP plan.
9. Migrate durable contracts into authoritative specs if anything new was learned, then move the WIP plan to `Specs\Plans\Done\`.

## Archived Payload

- `run-artifacts\full_after_commands_fileops_green_20260703.log`
- `run-artifacts\compare_decision_cache_pending_focused_after_full_red_20260703.log`
- `run-artifacts\msbuild_redsalamander_after_compare_pending_pin_spill_20260703.log`
- `run-artifacts\compare_decision_cache_pending_focused_after_pin_spill_20260703.log`
- `run-artifacts\compare_search_service_ntfs_seed_focused_after_full_red_20260703.log`
- `run-artifacts\compare_broad_after_pending_pin_spill_20260703.log`
- `run-artifacts\viewerpe_vlc_tab_focus_focused_after_full_red_20260703.log`
- `source-snapshots\RedSalamander_SelfTest_CompareDirectories_CompareDirectoriesEngine.SelfTest.Cases.RuntimeAndRemote.cpp`
- `source-snapshots\Tests_ViewerPETests_ViewerPETests.cpp`
- `source-snapshots\Plugins_ViewerVLC_ViewerVLC.h`
- `source-snapshots\Plugins_ViewerVLC_ViewerVLC.cpp`
- `source-snapshots\Tools_Tests_TestInventory.Tests.ps1`
- `source-snapshots\Specs_Testing_Testing_TestCoverage.md`
- `source-snapshots\Tests_README.md`
- `source-snapshots\Specs_Plans_WIP_Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md`
- `patches\focused-unstaged-diff.patch`
- `patches\focused-staged-diff.patch`
- `git-status-short.txt`
- `git-diff-stat.txt`
- `git-diff-cached-stat.txt`

## Do Not Forget

- Do not stage or copy the ignored recursive Full archive folder `Specs\TestRuns\4cb089111a23\Full\2026-07-03_202018_full_after_commands_fileops_green_red`; it is intentionally not authoritative.
- Do not weaken the ViewerVLC HUD focus contract just to make the test pass. Prove whether the HUD exists, is visible, has `WS_TABSTOP`, and is reachable by Tab from the video surface.
- Do not move the WarpDrive plan to Done until Full and both final FolderView perf matrices are green from the current HEAD.
