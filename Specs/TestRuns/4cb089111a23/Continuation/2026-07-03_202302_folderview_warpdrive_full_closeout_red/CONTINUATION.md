# FolderView WarpDrive Continuation - 2026-07-03 20:23

## Resume Pointer

Continue from this archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-03_202302_folderview_warpdrive_full_closeout_red\
```

Active branch: `codex/folderview-warpdrive`.

Active WIP plan:

```text
Specs\Plans\WIP\Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md
```

Do not move the plan to Done yet. Full-suite closeout is still red.

## What Changed Since The 18:14 Baton

- Root cause for the plugin-configuration S3 tab traversal red was the selftest predicate, not product tab order. `sendTab(...)` required `panelScrollPosY == 0` for every tab step, but focusing `Use virtual-hosted style addressing` legitimately scrolls the S3 settings panel. The wait predicate rejected the correct focus state, retried the debug tab advance, then sampled `Max keys per request`.
- Narrow patch in `RedSalamander\SelfTest\Commands\Commands.SelfTest.PluginConfig.cpp`: removed only the per-step `panelScrollPosY == 0` requirement. The test still asserts exact focus kind/label and DxUi-vs-legacy host counts; the initial top-of-panel setup remains intact.
- Debug build passed:
  `.build\codex-runs\msbuild_redsalamander_after_plugin_config_scroll_predicate_20260703.log`.
- Focused plugin-config tab traversal is green:
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_182138`.
- The first plugin-config local-window rerun at `2026-07-03_182158` is a command-line mistake: the filter file path was passed as the filter text, so zero cases ran. Ignore it as a product/test result.
- Correct plugin-config local-window rerun is green:
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_182308`.
- Exact `0..370` prefix is green:
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_183006`.
- Broad Commands is green:
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_184648` with 776 passed, 0 failed, 2 skipped.
- Standalone FileOps is green:
  `Specs\TestRuns\4cb089111a23\FileOps\2026-07-03_190019` with 102 passed, 0 failed, 20 skipped.

## Current Red Evidence

Full run:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2 *> .build\codex-runs\full_after_commands_fileops_green_20260703.log
```

Result: exit 1, 320 passed, 4 failed, 50 skipped, 374 total.

Preserved compact evidence:

- Runner stdout: `run-artifacts\full_after_commands_fileops_green_20260703.log`
- Top-level Full output files: `last-run-top-level\run-all-tests-results.json`, `last-run-top-level\trace.txt`, `last-run-top-level\ViewerPETests.output.log`, and the other executable output logs.
- Recursive `SelfTest\last_run\commands\work` fixture trees and raw `*.dmp` files are intentionally excluded; see `excluded-large-fixtures-and-dumps.txt`.

Failing suites/cases:

- CompareDirectories exit 1:
  - `decision_cache_eviction_budget_pending_wide_tree`: failed to drain pending subtree updates; `pendingSubdirUpdates=41`, `contentPendingCompares=540`, `contentQueueSize=536`, `contentInFlightSize=540`.
  - `search_service_sqlite_ntfs_traversal_seed_stays_degraded`: forced NTFS traversal-seed service did not finish startup warmup.
- ViewerPETests exit 1:
  - `TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly fresh harness run passes`
  - `TestViewerShellComboHostsLongRunOpenCloseStayStable fresh harness run passes`
- ToolsPesterTests exit 1:
  - `CompareDirectories static RunCase count drifted. Expected '229' but got '230'.`
  - `Coverage spec CompareDirectories count drifted from source. Expected '230' but got '229'.`

Everything else in the Full run was green: Commands, FileOperations, DxUiTests, FileSystemCurlTests, ViewerSqliteTests, MonitorTest, LocalizationTests, RedConfigureTests, PluginContractTests, SettingsSchemaTests, CrashHandlingTests, RedSalamanderMonitorEtwLatency, PerformanceTests2, and VcpkgMergeSynthetic.

## Next Commands

Run these sequentially. Do not run multiple `Run-AllTests.ps1` invocations in parallel; they share `SelfTest\last_run`.

```powershell
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter "decision_cache_eviction_budget_pending_wide_tree" -TimeoutMultiplier 2 *> .build\codex-runs\compare_decision_cache_pending_focused_after_full_red_20260703.log
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter "search_service_sqlite_ntfs_traversal_seed_stays_degraded" -TimeoutMultiplier 2 *> .build\codex-runs\compare_search_service_ntfs_seed_focused_after_full_red_20260703.log
```

Then inspect whether the Compare failures reproduce focused. If they pass focused, rerun a combined filter or broad Compare before changing product code.

ViewerPE focused repros:

```powershell
.\.build\x64\Debug\ViewerPETests.exe TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly *> .build\codex-runs\viewerpe_vlc_tab_focus_focused_after_full_red_20260703.log
.\.build\x64\Debug\ViewerPETests.exe TestViewerShellComboHostsLongRunOpenCloseStayStable *> .build\codex-runs\viewerpe_shell_combo_longrun_focused_after_full_red_20260703.log
```

Pester inventory drift:

```powershell
rg -n "CompareDirectories.*229|CompareDirectories.*230|static RunCase|Coverage spec|RunCase count" Tools\Tests Tests\README.md Specs\Testing RedSalamander\SelfTest\CompareDirectories
```

After focused fixes:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -TimeoutMultiplier 2
.\.build\x64\Debug\ViewerPETests.exe
pwsh -NoProfile -File .\Tools\Run-Pester.ps1
.\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2
```

Only after Full is green, rerun the final Debug and test-enabled Release FolderView perf matrices and cite the new archives before moving the WIP plan to `Specs\Plans\Done\`.

## Archive Contents

- `run-artifacts\`: exact build/test stdout logs and filter files from this leg.
- `last-run-top-level\`: compact Full-suite `SelfTest\last_run` files, excluding recursive work fixtures.
- `test-archive-summaries\`: top-level result/trace files copied from the key Commands/FileOps archives.
- `source-snapshots\`: current snapshots of PluginConfig, ViewerPE, TestInventory, Testing_TestCoverage, Tests README, and the WIP plan.
- `patches\unstaged-pluginconfig-and-plan.patch`: focused unstaged diff for the plugin-config predicate plus WIP plan once updated.
- `git-status-short.txt`, `git-diff-stat.txt`, `git-diff-cached-stat.txt`, `git-head.txt`, `git-branch.txt`.

