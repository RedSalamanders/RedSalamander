# I36 remediation follow-ups

Companion to [the remediation plan](../CodeReview_MasterMerges_2026-10-05.md) (I36, renumbered from I30 on 2026-10-11; the archived delivery record keeps I30). Non-normative record of what checking
each committed item turned up: owner actions, verification still owed, and Low items left open by the reviews.
Updated 2026-10-08 after commit `1ed166a1f`.

## Owner actions

1. **Recycle Bin leftovers.** Two test items from one failing-first experiment (run
   `20261007T150226Z-91352-4a264e9b4a7d4834bb3a6972f8c86d35`) are in the user's Recycle Bin: original names
   `rspurge_01DD566CDF00960B_000146CC.txt` and `rspurge_01DD566CDF00960B_000146CC_tree`, original folder
   `...\fileops\work\recyclebin-purge` of that run. Delete them from the Recycle Bin. Since `5ae046033` every recycle
   self-test purges its own items.
2. **One interactive window** (the owner chose a single focus-taking run at the end), covering:
   - the full FileOps suite, including `FileOps_AdmissionBaseline` (F7) and the `FileOps_SafetyRemediationContracts`
     telemetry witness, which failed on foreground changes while the desktop was in use;
   - Commands cases that the changes touch but that cannot run without focus:
     `file_action_selected_paths_identity_cleanup`, `file_action_selected_paths_recovery_bounds`,
     `cmd_pane_batchRename_recursive_collection_skips_symlink_loops`,
     `cmd_pane_batchRename_admission_worker_owned_cancellable`, `cmd_fileops_durable_store_safe_mechanics`,
     `cmd_fileops_move_breadcrumb_durable_lifecycle`, `cmd_pane_goToShortcutOrLinkTarget_junction_navigates_to_target`,
     `cmd_pane_executeOpen_junction_navigates_without_crash`, `cmd_pane_itemProperties_show_shortcut_and_reparse_targets`,
     `cmd_preferences_dialog_repeated_close_releases_hosts`, `cmd_pane_fileops_completed_group_and_navigation`,
     `cmd_pane_fileops_move_removal_focus_selects_next_survivor`,
     `cmd_pane_find_dialog_partial_completion_removes_only_known_sources`, and the Batch Rename set
     (`cmd_pane_batchRename_execution_engine_direct`, `..._partial_batch_failure_tracks_completed_rows`,
     `..._cancel_does_not_apply_remaining_rows`, `..._cancel_mid_batch_tracks_completed_rows`,
     `..._admission_rejects_cycle`, `..._window_executes_chain_rename`, `..._window_directory_chain_undo_plan`,
     `..._window_reports_execution_summary`, `..._execution_engine_large_independent_perf`);
   - the Commands suite-end Recycle Bin purge path, which only Commands exercises.
3. **One elevated run** for the test-owned VHDX path of F7 (`fc222f854`, `32202ee19`), never executed so far. Confirm
   the `Get-Disk` facts the pre-format proof expects for an attached fixed 128 MB VHDX: size 134217728, `Location`
   equal to the VHDX path, bus type "File Backed Virtual", partition style RAW. A mismatch fails safe (the disk is
   refused before formatting) but would skip the mounted routes on every elevated run.
4. **Decisions:**
   - Three frozen terminal-engine Pester tests create junctions with plain `New-Item` and are allowlisted because
     their bytes are pinned; their CI lanes run with `TEMP` outside a marked sandbox. Accept the documented exception,
     or redirect `TEMP` into a marked sandbox in `terminal-engine-gate0.yml` and `terminal-engine-round4-closeout.yml`.
   - A Native Move whose source vanished is reported "not removed by this item", so Retry is offered and fails again,
     and "Select retained" can list a missing path. Reporting "absent, not removed by us" needs a change to the
     shared receipt contract of every provider.

## Verification still owed

- After the next full Debug rebuild: `Get-RSPeTestSurface` on `.build\x64\Debug\RedSalamander.exe` must report the
  ENABLE_TESTS version marker added by `07f3167a9` (not yet observed on a real binary).
- No smoke launch of a tests-off Release build has been done.
- `R0fSmb_LoopbackRecycleNeverDestroys` skips when the loopback admin share is unreachable (common under a filtered
  token); it did run on the development machine.
- Gallery scenarios 136 and 137 changed with `a83359175`; re-baseline their hashes in
  `FileOperationsUiReview_2026-09-16/implementation/2026-10-01-v55-final` at the next gallery review.

## Behavior notes

- Ad hoc runs of link-creating Pester files and of `PluginContractTests`' FileSystem debug self-tests fail unless
  `TEMP`/`TMP` point inside a marked sandbox; `Run-AllTests` already does this (owner decision Q12).
- Interrupted-Move notices are no longer members of the Completed group: never hidden behind a badge, not counted in
  "Completed (N)" or its badges.
- The host treats `COPYENGINE_E_RECYCLE_FORCE_NUKE` as "destroyed, not recycled"; any provider that advertises
  Recycle must use it with that meaning (to be stated in `Plugins_VirtualFileSystem.md`).

## Low items left open by reviews

Tooling and tests:
- `9a7abe5a1`: a `-NonInteractive` run that executed some entries while others stayed pending still reports PASSED on
  suites without an evidence run; pending entries are printed but not recorded in `run-all-tests-results.json`; the
  inventory has no per-case activation classes, so no-activation FileOps cases cannot run unattended through
  `Test-Changes`.
- `d85927d0d`: the exposure source contract scans `Plugins/**/*.Tests.*` only, so plugin debug self-tests inside
  production translation units (beyond `RunDebugObjectBindingSelfTest`) are outside it.
- `29495bbaf`: the loopback-share contract finds the definition with a text heuristic.
- `5ae046033`: no test for the run-scope refusal of the Recycle Bin purge; the Commands suite-end drain only traces
  when its 30 s wait expires and re-implements bounded pumping instead of `PumpMessagesUntil`;
  `Core_SharedHelpers.md` lacks `PurgeRecycledTestItemsAtSuiteEnd`; the recycle journal lives in `FileSystem.dll`
  memory, so a module reload between recycle and purge would lose records.
- `251c042d2`: a read-only package folder ends empty but present (fails closed); the up-to-date MSIX fallback matches
  by identity anywhere under `AppPackages`; `build.ps1` records the `AppPackages` baseline before taking the
  packaging lease (pre-existing).
- `32202ee19`: a FileOps classification retry reuses the mounted VHDX and wipes the first failure's mounted witness
  (pre-existing).
- `--fileops-selftest --selftest-case` rejects comma-separated lists although `--help` says otherwise.
- `e8fa0be57`: the skip and cancel rounds check the task HRESULT loosely.

Product:
- `f67bf98c5`: batched `PreDeleteItem` and `PostDeleteItem` both match paths linearly per batch, and a refused entry is
  replayed through a Shell call that refuses again (no perf evidence).
- `30cd32412`: the BulkItems (pane Delete) folder-escalation test landed with the Q2 item; the pre-prompt size
  walk ignores Pause and shows no progress; a Failed folder after a partial recursive delete may already have lost
  descendants, which the docs should say.
- Q2 task-scoped escalation: the item text said "an item with unknown outcome still prompts". The implementation keeps
  the existing fail-closed rule instead: an indeterminate Recycle never prompts, is never offered permanent deletion
  and is never covered by the grant (`Phase10_RecycleEscalationTaskScopedGrant` steps 7-9). Owner confirmation
  pending. The pane Delete (BulkItems) path now goes on past its provider call's first failure after any resolved
  escalation (Skip included), as the per-item executor does.
- `260c9189a`: a per-item setup failure before any dispatch still collapses the task HRESULT to
  `ERROR_PARTIAL_COPY` for multi-root tasks.
- `f290374e9`: a rename merge that only removed the emptied root reports publication Published while
  `bridgePublication` says NotPublished.
- `e8fa0be57`: the spec sentence on invoked-but-rejected Batch Rename rows needs the Canceled and Skipped cases; never
  run rows show the system text of `ERROR_OPERATION_ABORTED` in Issues (fits the localized primary-text item).

## Incidents

- 2026-10-07 about 21:30: drive Z: filled to 0 bytes (mostly data outside this project). One Debug relink failed and
  one `git add` failed and was retried; a stale empty `.git/index.lock` from that window was removed. All objects of
  the run's commits were verified present and the Debug build was rebuilt cleanly after the owner freed space.
