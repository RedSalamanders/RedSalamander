# Three-Day Review / Folder View — Continuation Baton

Status: **complete on 2026-07-12**. The broad-suite blockers were resolved and the final repository-wide gate is green.

## Objective

Finish broad-suite closeout for the user-requested last-three-days code review, high-to-low remediation, and the Folder View deletion regression fix.

## Completed

- The Folder View regression is diagnosed, fixed, and focused-tested. Successful File Operations dismissal could race the settings watcher's classification of the application's own atomic save; the resulting plug-in refresh cleared both pane models, and restoration did not re-enumerate them. The remediation fences self-writes and process lifetime, preserves raw provider path/context, qualifies once, and explicitly re-enumerates both panes. Focused coverage passed 4/4 and repeated 20/20.
- The reviewed findings and remediation ledger are closed in:
  - `../Done/CodeReview_Last3Days_IndependentFindings_2026-07-11.md`
  - `../Done/CodeReview_Last3Days_Remediation_2026-07-11.md`
- ViewerSqlite, ViewerWeb, ViewerPE, and ViewerText focused builds/tests passed. Source contracts are 132/132; Tools inventory is 253 and its inventory contract passed 5/5.
- The stale Compare self-test NTFS-only precondition was corrected to accept the production-supported NTFS or ReFS root. Exact 1/1, capability matrix 2 pass/1 legitimate skip, and repeat 5/5 passed.

## Resolved blockers

1. Commands case `pane_view_options_preview_uses_configured_embedded_viewer_and_preserves_focus` was isolated to outstanding ViewerVLC startup jobs contaminating a gated retirement-cleanup assertion. The debug snapshot now exposes pending load work and the case waits for unrelated work to settle before measuring retirement. Focused repeat proof and broad Commands are green.
2. The wide-tree Compare case passed in isolation and in the final broad/full order. The prior apparent stop was concurrent-run contention, not an independent product hang.
3. Broad-order follow-ups fixed recently-missing selected-item continuity across delayed chained rename hints and scoped the NavigationView history test to its owned DxUi popup.

## Evidence and exact handoff

The authoritative detailed checkpoint, including root-cause narrative, validation logs/run IDs, exact resume commands, and cautions, is:

`Specs/TestRuns/4cb089111a23/Continuation/2026-07-11_2315_review_checkpoint/README.md`

Earlier Full-run traces and partial results are preserved in:

`Specs/TestRuns/4cb089111a23/Continuation/2026-07-11_2249_full_commands_preview_pane_hang/`

Readable full live-process dump (intentionally outside tracked specs because it is 1.09 GB):

`.build/codex-runs/RedSalamander-hang-analysis2-p98512.dmp`

The dump is 1,088,958,364 bytes and has a verified `MDMP` header. The older `RedSalamander-hang-20260711-2249-p37660.dmp` is incomplete and must not be used.

## Closeout result

`Tools/Run-AllTests.ps1 -Suite Full` completed with 1,132 passed, 0 failed, and 53 documented skips. Compare was 218/0/31, Commands 797/0/2, and File Operations 103/0/20. The bounded final archive is `Specs/TestRuns/4cb089111a23/Continuation/2026-07-12_205636_farsight_full_green/`.

## Safety

The shared worktree is extensively dirty with concurrent changes. Preserve everything unrelated. No commit, staging, branch switch, reset, or cleanup was performed for this checkpoint.
