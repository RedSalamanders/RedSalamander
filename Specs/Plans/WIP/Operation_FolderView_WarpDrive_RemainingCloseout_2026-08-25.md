# FolderView WarpDrive Remaining Closeout

> **NON-NORMATIVE ACTIVE PLAN.** Compact remainder of the archived master `Specs/Plans/Done/Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md`. Tasks 0–9 are done. Task 10 Steps 2–3 are measured no-ops (gate not met — do not implement DXGI scroll-rects or virtualization V2). Thumbnail enrichment is **not** this plan: `FolderView_ThumbnailBackgroundEnrichmentFollowup_2026-07-04.md` (I1). Retired continuation batons live in Done and must not be resumed.

## Status

- **State:** ACTIVE
- **Priority:** P2 FolderView perf closeout
- **Planned at:** `56db9af9d0a19b4b3f375e16989591f29e014a6c`
- **Ownership boundary:** Task 10 Step 4 verification-or-waiver and Task 11 spec/Full closeout only. No new FolderView rendering architecture.
- **Drift check:** `git diff 56db9af9d0a19b4b3f375e16989591f29e014a6c..HEAD -- Specs/UI/UI_FolderView.md Specs/Testing/Testing_PerformanceValidation.md Specs/Testing/Testing_TestCoverage.md Specs/Testing/FolderViewPerfBudgets.json5 RedSalamander/FolderView.Rendering.cpp RedSalamander/FolderView.Interaction.cpp`
- **Authoritative specs:** `Specs/UI/UI_FolderView.md`, `Specs/Testing/Testing_PerformanceValidation.md`, `Specs/Testing/Testing_TestCoverage.md`

## Remaining outcomes

1. **Task 10 Step 4 — verify or explicitly waive Full-as-Task-10-gate.** Run the full FolderView perf matrix in Debug and test-enabled Release, then `.\Tools\Run-AllTests.ps1 -Suite Full`, **or** record an explicit maintainer waiver that Task 11 Step 3 is the single Full/perf gate (do not require two Full runs). Do not reopen dirty-region V2, DXGI scroll-rects, or item-identity virtualization; those gates were not met.

2. **Task 11 Step 1 — spec remainder only.** Most Task 11 Step 1 contracts already live in `Specs/UI/UI_FolderView.md` (cached-only visible thumbnails, device-loss recovery, draw-loop brush reuse, async paste-shortcut generation gating, `folder.refresh.*` and icon-pipeline metrics). Background enrichment stays out of this plan (I1). Do not copy diary into the spec.

3. **Task 11 Step 2 — testing-spec remainder.** Confirm sample-count, Release/test-enabled evidence, budget, scale/cold/slow fixture, and `Specs/TestRuns/` archive rules in testing specs; coordinate leftover enforcement with I5 rather than duplicating a second measurement contract.

4. **Task 11 Step 3 — closeout evidence.** After operator-authorized Full and the current-HEAD Debug plus test-enabled Release FolderView perf matrix are green and cited, merge any last durable sentences into the specs above and move this remainder to `Specs/Plans/Done/`.

## Verification

```powershell
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
git diff --check
```

Full and the FolderView perf matrix run only when a maintainer authorizes them as Task 10 Step 4 / Task 11 Step 3. Do not treat June/July pause-diary Commands flakes as current next actions.

## Done criteria

- Task 10 Step 4 is verified or explicitly waived against Task 11 Step 3.
- Durable FolderView/testing contracts for completed WarpDrive work are in current specs (enrichment remains I1).
- Cited Full and perf-matrix evidence exist, or a recorded waiver names the substitute gate.
- This remainder is moved to Done; the 2026-06-28 archive stays frozen.

## STOP conditions

Stop if Tasks 0–9 have been unchecked in the Done archive (product regression — do not reactivate retired batons). Stop if new evidence beats the Task 10 Step 2/3 risk threshold and would require DXGI scroll-rects or V2; that needs a new ACTIVE plan, not this closeout file. Stop if Full is red for reasons outside FolderView perf closeout and the failure is claimed as WarpDrive product work without a named owner.
