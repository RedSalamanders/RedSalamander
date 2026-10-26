# FolderView WarpDrive continuation baton - Preferences root/detail red

Archive created: 2026-07-02 10:37 local time.

Branch: `codex/folderview-warpdrive`.

This supersedes `Specs/TestRuns/4cb089111a23/Continuation/2026-07-02_102745_folderview_warpdrive_prefs_commands_baton/` as the immediate resume pointer.

## Current state

- No product/test patch has been applied after this new investigation yet.
- The worktree remains dirty and heavily staged from the broader WarpDrive implementation. Do not reset or revert unrelated changes.
- The latest active blocker is still broad-suite Preferences/DxUi order sensitivity, not the original FolderView perf gates.
- Do not move `Specs/Plans/WIP/Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md` to Done until broad Commands, Full, and final Debug plus test-enabled Release FolderView perf matrices are green and cited.

## Fresh evidence archived here

- `run-artifacts/prefs_plugins_fourcase_rootleak_probe_20260702.log`
  - Passed.
  - Cases: Plugins UIA grid, Plugins custom-paths UIA grid, Plugins long-run scroll, Plugins custom-paths long-run scroll.
- `run-artifacts/prefs_dialog_family_rootleak_repro_before_fix_20260702.log`
  - Failed.
  - Command: `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -CaseFilter 'cmd_preferences_dialog_' -TimeoutMultiplier 2.0`
  - Result: 14 passed / 1 failed / 158 skipped.
  - Failed case: `cmd_preferences_dialog_viewers_live_search_dx_interaction`.
  - Failure reason: bare `failed`; the test returned false without setting a diagnostic failure string.
- Volatile last-run files copied:
  - `run-artifacts/prefs_dialog_family_rootleak_repro_run-all-tests-results.json`
  - `run-artifacts/prefs_dialog_family_rootleak_repro_commands-results.json`
  - `run-artifacts/prefs_dialog_family_rootleak_repro_commands-trace.txt`
  - `run-artifacts/prefs_dialog_family_rootleak_repro_trace.txt`
- Repo state copied:
  - `git-status-short.txt`
  - `git-diff-stat.txt`
  - `git-diff-cached-stat.txt`
  - `git-diff-name-status.txt`
  - `git-diff-cached-name-status.txt`
  - `patches/staged-diff.patch`
  - `patches/unstaged-diff.patch`
- Source snapshots copied:
  - `source-snapshots/Preferences.Dialog.cpp`
  - `source-snapshots/Preferences.Plugins.cpp`
  - `source-snapshots/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp`
  - `source-snapshots/Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp`

## Investigation notes

### Plugins root/detail state leak

Evidence from the previous 10:27 baton showed a broad Commands failure at `cmd_preferences_dialog_plugins_page_exposes_live_uia_grid_selection` with:

- `title='S3'`
- `pluginItemSelected=true`
- `pluginsDetailsActive=true`

Focused and 14-case predecessor reruns passed, which points at retained Preferences state or navigation order rather than a simple deterministic functional failure.

The likely product-contract bug is in `PreferencesDialog::SelectCategory(...)`:

- `RedSalamander/Preferences.Dialog.cpp` around `SelectCategory` clears:
  - `state.pluginsSelectedPlugin`
  - `state.pluginsSelectedPluginId`
  - `state.pluginsDetailsActive`
- It does **not** clear `state.pluginsRetainedSelectedPluginId`.
- The tree delegate category-root path does clear `pluginsRetainedSelectedPluginId`, so `DebugSelectPreferencesCategory(kPrefCategoryPlugins)` does not currently match real category-root selection semantics.
- `RedSalamander/Preferences.Plugins.cpp` restores selected plugin cache from `pluginsRetainedSelectedPluginId` when the Plugins page refreshes, so stale retained plugin state can make root-selection tests observe a plugin row/detail state.

Candidate product fix:

```cpp
state.pluginsRetainedSelectedPluginId.clear();
```

Add it in `SelectCategory(...)` next to the other Plugins root-state clears, then verify the Plugins four-case probe and the Preferences-family run.

### Viewers live-search silent failure

The fresh Preferences-family rerun failed earlier at `cmd_preferences_dialog_viewers_live_search_dx_interaction`.

The trace says:

```text
Case: cmd_preferences_dialog_viewers_live_search_dx_interaction
Case returned: cmd_preferences_dialog_viewers_live_search_dx_interaction ok=no failed=no skipped=no
CommandsSelfTest: FAIL
failed
```

That means the test returned `false` without calling `state.Require(...)` or otherwise setting a useful failure reason.

Code to inspect first:

- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp`
- `TestPreferencesDialogViewersLiveSearchDxInteraction(...)`
- Its local `navigateToViewersPage` helper currently uses relative keyboard navigation:
  - focus category tree host
  - `VK_HOME`
  - `VK_DOWN` twice
  - wait for Viewers snapshot
- If navigation fails, the test returns false silently.

Candidate selftest hardening:

- Prefer `DebugSelectPreferencesCategory(kPrefCategoryViewers)` for this setup, matching newer deterministic Preferences helpers elsewhere.
- If keeping any keyboard fallback, wrap it in `state.Require(...)` with a formatted snapshot that includes current category, title, selected plugin flags, row counts, focus HWND, and page-host resize failure count.

## Next safe implementation slice

1. Patch `PreferencesDialog::SelectCategory(...)` to clear `pluginsRetainedSelectedPluginId` when selecting a category root.
2. Patch `cmd_preferences_dialog_viewers_live_search_dx_interaction` setup so Viewers navigation is deterministic and failures are diagnostic, not bare `failed`.
3. Rebuild Debug:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
```

4. Run focused and reduced checks:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -CaseFilter 'cmd_preferences_dialog_viewers_live_search_dx_interaction' -TimeoutMultiplier 2.0
```

```powershell
$filter = 'cmd_preferences_dialog_plugins_page_exposes_live_uia_grid_selection,cmd_preferences_dialog_plugins_custom_paths_page_exposes_live_uia_grid_selection,cmd_preferences_dialog_plugins_long_run_list_scrolling_stays_bounded,cmd_preferences_dialog_plugins_custom_paths_long_run_list_scrolling_stays_bounded'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -CaseFilter $filter -TimeoutMultiplier 2.0
```

5. Re-run the Preferences family fail-fast:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -CaseFilter 'cmd_preferences_dialog_' -TimeoutMultiplier 2.0
```

6. Only after that goes green, return to broad Commands fail-fast. Do not run multiple `Run-AllTests.ps1` invocations in parallel; they share `SelfTest\last_run`.

## Remaining before Done

- Broad Commands green on current branch.
- Full suite green, or any non-green result fully isolated and fixed.
- Final Debug FolderView perf matrix from current HEAD.
- Final test-enabled Release FolderView perf matrix from current HEAD.
- WIP plan updated with final evidence.
- Durable behavior/spec notes migrated out of WIP where applicable.
- Move the WIP plan to `Specs/Plans/Done/` only after all of the above.
