# FolderView WarpDrive Continuation Baton - Compare Cache Crash

Created: 2026-07-03 14:26 Europe/Paris
Repo: `Z:\src\RedSalamander`
Branch: `codex/folderview-warpdrive`
HEAD: `275c040348a3994d45c6aef340333a35a0490cb6`
Origin status at archive time: `origin/master...HEAD = 4 behind / 25 ahead`

## Start Here

The next blocker is a CompareDirectories access violation, not the older PerfJsonl thumbnail crash.

Latest failing command:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -TimeoutMultiplier 2
```

Evidence:

- `run-artifacts\compare_full_after_quiet_drain_20260703.log`: `CompareDirectories` child exited `-1073741819`; no `compare\results.json` was written.
- `run-artifacts\last_run_run-all-tests-results.compare-crash.json`: confirms `suite=Compare`, `passed=0`, `failed=0`, child exit `-1073741819`.
- `crash-artifacts\RedSalamander-20260703-142049-p98100.txt`: access violation in `CompareDirectoriesSession::ComputeDecisionForFolder`, `CompareDirectoriesEngine.cpp:3551`, inside `std::unordered_map<ContentCompareKey, bool>::find`.
- `crash-artifacts\RedSalamander-20260703-142049-p98100.dmp`: paired dump for the same crash.

Likely investigation line: content-decision-cache concurrency/lifetime. The crash happened on a scan worker, while probing the decision cache from `ComputeDecisionForFolder`; inspect synchronization around the cache and any recently widened Full/Compare ordering that lets content comparison workers mutate it concurrently.

## Current State

Current WIP plan:

- `Specs\Plans\WIP\Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md`
- The plan roll-up now has a 2026-07-03 Task 11 entry pointing to this baton and naming the crash.

Already green on the current worktree:

- Debug RedSalamander build after the Compare/Commands wait patches:
  - `run-artifacts\msbuild_redsalamander_after_compare_navedit_waits_20260703.log`
- Focused CompareDirectories pending-wide-tree case after quiet-drain helper:
  - `run-artifacts\compare_pending_wide_tree_focused_after_quiet_drain_20260703.log`
  - case: `decision_cache_eviction_budget_pending_wide_tree`
- Focused Commands navigation-edit clipboard case after host/input-open predicate:
  - `run-artifacts\commands_clipboard_nav_edit_focused_after_host_open_predicate_20260703.log`
  - case: `cmd_pane_clipboardPaste_ignores_unfocused_navigation_edit`
- Full CompareDirectories before the quiet-drain patch:
  - `run-artifacts\compare_full_after_pending_window_green_20260703.log`
  - 207 passed / 0 failed / 30 skipped.
- Broad Commands before the latest Full rerun:
  - `run-artifacts\commands_broad_after_nav_pref_waits_20260703.log`
  - 776 passed / 0 failed / 2 skipped.
- ViewerPE direct and Monitor ETW direct were green in the previous baton:
  - `Specs\TestRuns\4cb089111a23\Continuation\2026-07-03_132317_folderview_warpdrive_fullsuite_compare_monitor_baton\CONTINUATION.md`

Latest proper Full-with-build state before the focused fixes:

- `run-artifacts\full_with_build_after_compare_monitor_green_20260703.log`
- `.\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2`
- 1097 passed / 2 failed / 52 skipped.
- Red cases:
  - CompareDirectories: `decision_cache_eviction_budget_pending_wide_tree`, failed to drain pending subtree updates.
  - Commands: `cmd_pane_clipboardPaste_ignores_unfocused_navigation_edit`, navigation edit was open but the old predicate required a transient `PathEdit` focus target.

## Code Changes Since Prior Baton

Focused patches added after the previous 13:23 baton:

- `RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.cpp`
  - Added `DrainPendingSubdirUpdatesUntilQuiet(...)`, requiring two quiet samples of no pending subtree updates and no budgeted flush work.
- `RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.Cases.RuntimeAndRemote.cpp`
  - Switched `decision_cache_eviction_budget_pending_wide_tree` to the quiet-drain helper with scaled 10s timeout and richer failure stats.
- `RedSalamander\SelfTest\Commands\Commands.SelfTest.ShellCommands.cpp`
  - Changed `TestPaneClipboardPasteIgnoresUnfocusedNavigationEdit` to wait for navigation edit host/input HWNDs instead of requiring transient focus target `PathEdit`.
- `Specs\Plans\WIP\Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md`
  - Updated the Task 11 roll-up with current green gates and the Compare cache crash.

Patch snapshots:

- `patches\unstaged-code-and-spec-no-testruns.patch`
- `patches\staged-code-and-spec-no-testruns.patch`

Source snapshots:

- `source-snapshots\RedSalamander_SelfTest_CompareDirectories_CompareDirectoriesEngine.SelfTest.cpp`
- `source-snapshots\RedSalamander_SelfTest_CompareDirectories_CompareDirectoriesEngine.SelfTest.Cases.RuntimeAndRemote.cpp`
- `source-snapshots\RedSalamander_SelfTest_Commands_Commands.SelfTest.ShellCommands.cpp`
- `source-snapshots\Specs_Plans_WIP_Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md`
- `source-snapshots\.agents_skills_red-salamander-crash-forensics_SKILL.md`

## Resume Plan

1. Use the local crash skill before changing code:

   ```powershell
   Get-Content .\.agents\skills\red-salamander-crash-forensics\SKILL.md
   ```

2. Inspect `CompareDirectoriesEngine.cpp` around `ComputeDecisionForFolder`, especially line 3551 and the content-decision-cache container ownership/locking.

3. Prove the cache mutation/read path. Search for every access to the content decision cache and confirm all readers/writers share the same lock or are otherwise immutable during scan-worker reads.

4. Fix the crash first, then rerun:

   ```powershell
   .\build.ps1 -ProjectName RedSalamander -Configuration Debug
   .\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter "decision_cache_eviction_budget_pending_wide_tree" -TimeoutMultiplier 2
   .\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -TimeoutMultiplier 2
   ```

5. After Compare is green, rerun the Commands clipboard/navigation-edit window and broad Commands:

   ```powershell
   .\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_pane_clipboardPaste_path_change_clears_stale_overlay,cmd_pane_clipboardPaste_ignores_unfocused_navigation_edit,Commands_FileOpsClipboardPasteUsesHostQueue,Commands_FileOpsFolderPickerMoveUsesHostQueue,Commands_FileOpsMissingCallbackRejectsDirectFallback" -TimeoutMultiplier 2
   .\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2
   ```

6. Then rerun proper Full with build:

   ```powershell
   .\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2
   ```

7. Only after Full is green, rerun and cite final Debug plus test-enabled Release `folderView_perf_` matrices, update durable specs if any new contract was discovered, then move the WIP plan to `Specs\Plans\Done\`.

## Archive Contents

Use `archive-manifest.txt` for the complete file list. The key directories are:

- `run-artifacts\`: latest build/test logs and volatile `last_run` JSON.
- `crash-artifacts\`: latest RedSalamander crash sidecar/dump and crash-folder inventory.
- `patches\`: staged and unstaged non-TestRuns patches at archive time.
- `source-snapshots\`: relevant source/spec snapshots.

Do not treat the plan as Done from this baton. The current closeout status is still WIP: Compare crash, then Commands window/broad Commands recheck, then Full, then final perf matrices.
