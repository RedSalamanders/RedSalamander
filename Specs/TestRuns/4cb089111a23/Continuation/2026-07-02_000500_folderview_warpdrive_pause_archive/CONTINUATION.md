# FolderView WarpDrive pause archive - 2026-07-02 00:05

## Resume target

- Workspace: `Z:\src\RedSalamander`
- Branch: `codex/folderview-warpdrive`
- HEAD at archive time: `275c040348a3994d45c6aef340333a35a0490cb6`
- `origin/master...HEAD`: `0 25`
- No rebase was needed at archive time.

## Current state

- The worktree has a large staged WIP for `Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md`.
- Treat staged changes as valuable WIP. Do not reset, checkout, or discard them.
- Three original WIP baton files remain untracked in `Specs/Plans/WIP/`; snapshots are archived under `untracked-wip-snapshots/`.
- The active plan is not ready for Done. Task 10 and Task 11 remain open, and the current blocker is a broad Commands crash.

## Latest verified evidence

Copied from the previous archive:

- `Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_223000_folderview_warpdrive_origin_commands_recheck`

Known green evidence from that archive:

- Clean Debug build: `.build\logs\msbuild-20260701_222239_273.log`
- Focused Commands green:
  - `folderView_render_device_loss_recovers`
  - `folderView_dpi_change_repaints_both_panes`
  - `cmd_pane_navigationView_edit_mode_survives_external_refresh`
  - `cmd_pane_filter_prompt_live_dx_interaction`
- Reduced prompt slice `602..627` passed `26/0/0`.

Known red evidence:

- Broad Commands run still crashes.
- Wrapper exit: `1`
- App exit: `-1073741819` (`0xC0000005`)
- WER module: `textinputframework.dll`
- Fault offset: `0x000000000006a419`
- Trace stops during `cmd_pane_filter_prompt_live_dx_interaction` after:
  - `pane_filter_prompt_live_dx: cycle 'live Cancel interaction' worker-done`
  - before case return / before `joined`
- This is timing/order dependent: focused filter prompt and local prompt slice pass.

## Archived files in this folder

- `staged-diff.patch` - complete staged patch at pause time.
- `unstaged-diff.patch` - unstaged diff at pause time.
- `git-status-short.txt`, `git-branch.txt`, `git-head.txt`, `origin-master-ahead-behind.txt`
- `git-diff-cached-stat.txt`, `git-diff-stat.txt`
- `git-staged-name-only.txt`, `git-unstaged-name-only.txt`, `git-untracked-name-only.txt`
- `Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.snapshot.md`
- `untracked-wip-snapshots\*.md`
- Copied crash/build/test evidence from the latest prior archive:
  - `commands-broad-after-repaint-assertion-run-all-tests-results.json`
  - `commands-broad-crash-commands-trace.txt`
  - `commands-broad-crash-last-run-trace.txt`
  - `windows-error-reporting-recent.txt`
  - `application-error-redsalamander-recent.txt`
  - `commands-prompt-slice-602-627-run-all-tests-results.json`
  - `prompt-slice-602-627.filter.txt`
  - `msbuild-20260701_222239_273.log`
- `process-snapshot-redsalamander.txt`

## Next work

1. Inspect TSF/native-text lifecycle in:
   - `Common\DxUi\DxUi.TextStoreACP.cpp`
   - `Common\DxUi\DxUi.NativeTextInput.cpp`
   - `Common\DxUi\DxUi.WindowHost.cpp`
2. Inspect prompt teardown and modal worker/join boundaries in:
   - `RedSalamander\FolderWindow.FileSystem.cpp`
   - `RedSalamander\SelfTest\Commands\Commands.SelfTest.Dialogs.cpp`
3. Build a tighter non-traced replay that includes the preceding FileOps/issues/navigation shell cases plus the prompt cases. The passing `602..627` slice is too narrow.
4. If the reduced replay crashes, rerun the exact slice with `REDSALAMANDER_DXUI_TEXTINPUT_TRACE_FILE`, but treat a traced pass as timing perturbation, not proof.
5. Patch only after the lifetime/root-cause evidence is clear.
6. Rebuild Debug and rerun:
   - the four focused green cases above,
   - the new reduced red/green replay,
   - broad `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2.0`.
7. Only after broad Commands is green should the plan continue toward Task 10 and Task 11 closeout gates: Release/test-enabled matrix, perf evidence, durable spec migration, full suite, and move to Done.

## Guardrails

- Do not move the WIP plan to Done yet.
- Do not claim the implementation is ready until broad Commands no longer crashes and the remaining plan gates are satisfied.
- Do not discard staged files; the staged patch is the working continuation state.
