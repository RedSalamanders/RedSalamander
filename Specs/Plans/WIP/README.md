# Specs/Plans/WIP - Index & Ranked Queues

Last WIP triage update: **2026-07-16** after Lighthouse closeout reconciliation.

This README is the routing index for unfinished plans. It is not the source of truth for task details:
open the named plan before executing any item.

## Rules

- One owner per requirement. When a plan says `ROUTED to <owner>`, execute the owner, not the routed
  copy.
- Do not execute archived ledgers in `../Done/` as parallel WIP queues. They are reference evidence.
- The `HOLD-*` rows below are not available as normal next picks. Three are coordination-sensitive
  FolderView/DxUi ownership rows, and one is human-managed scratch state.
- Completed WIP plans must move to `Specs/Plans/Done/`, and durable behavior must be merged into the
  authoritative spec under `Specs/<Domain>/` before closeout.

## Current Audit Refresh

`Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15.md` is the current-head,
whole-repository defect and architecture report. It directly owns findings that had no prior owner and routes
the rest without creating competing branches. Lighthouse is complete and archived under `Specs/Plans/Done/`;
the first executable Observatory-owned slice is now I1 Track 0, restoring a trustworthy gate.

## Best Next Pick

If the current goal is remediation or stabilization, continue with **I1 Observatory Track 0**: classify the
remaining current-suite failures and consolidate inventory truth. Lighthouse, Firebreak, Causeway, Farsight,
and the test-suite stabilization plan are completed references; do not reopen their routed rows.

If the current goal is product work instead of remediation, start with **N1 Google Drive Phase 2**:
interactive OAuth PKCE sign-in plus read-stream/download support.

## Improvement / Remediation Queue

Ranked by impact, urgency, dependency value, and whether the work is already owned elsewhere.

| Rank | File | Why this rank | Next action / routing |
|------|------|---------------|-----------------------|
| I1 | `Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15.md` | Whole-repository refresh plus verified seven-day supplement: 82 findings, ordered by release, security, data-loss, correctness, architecture, and simplification risk. Existing owners remain authoritative for routed rows. | Execute **Track 0** against the current landed test set to classify the remaining gate failures and replace duplicated inventory truth; then take the stop-line tracks in order. |
| I2 | `Operation_FileOperations_FaultInjectionRatchet_2026-07-05.md` | Enables RED-able coverage for bridge and cloud data-safety gaps; useful before several FileOps fixes. | **FIR-1 / FG-A1**: destination-side bridge `GetSize` fault injection through the bridge IO decorator. |
| I3 | `Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md` | Cloud parity/perf proof queue; overlaps the coverage seam needed by FIR. | **FG-A1**: destination-side `GetSize` fault-injection hook. Coordinate with I2 so only one seam is built. |
| I4 | `Fix-to AWS-S3-crash.md` | Crash mitigation is in code, but runtime plugin refresh unload remains unproven. | **CL-1/CL-2**: run and archive the ASan heavy-S3 recipe including runtime plugin refresh, or close the unload gap in code. |
| I5 | `Operation_IronLedger_FolderViewDataIntegrityDropClipboard_2026-06-28.md` | FolderView drag/drop/clipboard data-integrity backlog with direct wrong-target/data-loss risks. | **Task 1**: drop-point plumbing and self/same-parent/descendant rejection in `FolderView::PerformDrop`. |
| I6 | `Operation_PerfMeasurementContract_2026-07-06.md` | Raises the quality bar for every perf-sensitive plan; not a runtime bug, but it prevents weak closeouts. | **Task 1**: install the Perf Measurement Record template in authoritative specs and add representative records. |
| I7 | `UI_FileOperationsPopupUxRefinementPlan_2026-07-07.md` | User-requested progress-popup clarity work; A-tier, queued-task Start now, footer bulk pause/resume, B4 completed grouping, and both review-remediation passes are implemented, but workflow/accessibility polish keeps the plan in WIP. | Continue remaining backlog: B1 reorder, B4 completed-group animation, B5 keyboard/UIA, C5 Keep Both engine support, C7 shared DxUi graph closeout. |
| I8 | `Code_PluginImprovementPlan.md` | Small remaining plugin lifecycle queue; one item was wrongly marked done. | **4.1**: coordinate `CurlEasyPool` drain and `curl_global_cleanup` at the plugin quiet point. |
| I9 | `FileSystem_CrossFsBridgeImplementationAndPerformanceRedesign_2026-07-07.md` | Design-level performance rethink for the cross-FS bridge; valuable after the immediate Causeway fixes. | Use as a dedicated design/perf slice; start with R-1 depth-N pump only with perf evidence. |
| I10 | `Win32_Inventory.md` | Narrow RAII/lifetime/DPI audit with bounded fixes. | **CHK-1**: replace `DestroyWindow(_hWnd.get())` on `wil::unique_hwnd` owners with reset/transfer-safe close. |
| I11 | `Operation_CommandsSelfTestInputIsolation_2026-06-24.md` | CI5-R implementation, source contract, build, and five affected GUI cases are green. | Run the broad Commands closeout suite, then move the plan to Done if green. |
| I12 | `Operation_Tailwind_FolderViewWarpDrivePreMergeReviewRemediation_2026-06-30.md` | Residual branch self-review ledger; many items are done or routed, so it is not the best primary queue. | Reconcile freshness first; execute only remaining non-routed items after WarpDrive/DxUi owners confirm no overlap. |
| I13 | `Core_HResultStatusFormattingCleanup_2026-06-19.md` | Low-risk cleanup and localization consistency, but low urgency. | **WI-1**: inline `FormatHResult` and `FormatStatusText`; coordinate with hot FolderView files. |
| HOLD-I14 | `Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md` | Active/coordination-sensitive closeout plan. | Do not pick unless the WarpDrive owner hands it off. |
| HOLD-I15 | `FolderView_WarpDrive_ContinuationBaton_2026-06-29.md` | Resume state for WarpDrive closeout and suite proof. | Do not pick unless the WarpDrive owner hands it off. |
| HOLD-I16 | `DxUi_Uia_ContinuationBaton_2026-06-29.md` | Active/coordination-sensitive UIA snapshot and commit-split baton. | Coordinate before touching DxUi/UIA snapshot coalescing or related routed Tailwind/Granite items. |
| HOLD-I17 | `Notes_Scratch.md` | Human-managed scratch file, not an executable agent queue. | Keep out; owner can delete or reconcile manually. |

## Deferred Decision Review Queue

These operations preserve previously rejected or policy-sensitive scenarios for an explicit maintainer decision.
They are not implementation queues; approved work must be routed to a separate executable WIP owner.

| Rank | File | Why it is retained | Next action / routing |
|------|------|--------------------|-----------------------|
| D1 | `Operation_Parallax_DeferredArchitectureSecurityAndSimplificationDecisionReview_2026-07-14.md` | Preserves twelve Lighthouse rejection/policy scenarios with explanations, evidence requirements, reopen triggers, and decision options. | Review PAR-1 through PAR-12 when the relevant security, product, or architecture policy is ready for decision. Record one outcome per scenario and route approved implementation separately. |

## New Feature Queue

Ranked by concrete implementation readiness and product value. These are user-visible additions or
feature-roadmap plans, not primarily bug-fix/remediation queues.

| Rank | File | Why this rank | Next action / routing |
|------|------|---------------|-----------------------|
| N1 | `FileSystem_GoogleDrivePluginPlan.md` | Concrete new filesystem plugin; enumeration landed, but no in-product first sign-in or file reads yet. | **Phase 2**: interactive OAuth PKCE sign-in plus read-stream/download support, using Microsoft Drive as the local model where applicable. |
| N2 | `FolderView_ThumbnailBackgroundEnrichmentFollowup_2026-07-04.md` | User-visible completion of the cached-only thumbnail decision; restores richer cold-cache thumbnails without blocking first paint. | Implement bounded background enrichment for local-shell-backed files with generation checks, metrics, and perf evidence. |
| N3 | `Terminal_EmbeddedConPtyLibGhosttyVt_IntegrationPlan_2026-07-13.md` | Implementation-ready embedded terminal architecture with explicit ConPTY, VT, lifecycle, security, accessibility, and perf gates. Large scope makes it less suitable than N1/N2 for the immediate next slice. | Start only as a dedicated branch from its Phase 0 capability ledger and dependency/license gates; do not mix it with FolderView or DxUi hold work. |
| N4 | `Product_WhimFilesGapAnalysisAndImprovementPlan_2026-07-08.md` | Product roadmap/gap analysis, not an implementation-ready slice. | Convert the highest-value gaps into separate executable plans before coding. Start with G1 live multi-dimensional filtering or G2 undo if product direction is approved. |

## Archived Reference Ledgers

These were consolidated or completed and should stay in `Specs/Plans/Done/` unless fresh evidence
reopens a specific item:

- `Operation_Crosscut_CompareDirectoriesRemediation_Handoff_2026-06-22.md`
- `UI_BatchRenameWindowPlan_2026-06-10.md`
- `Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
- `Operation_Granite_ContinuationBaton_2026-07-05.md`
- `Operation_Bedrock_ThreeDayDiffReviewRemediation_2026-07-06.md`
- `Operation_Evergreen_ThreeDayDiffReviewRemediation_2026-07-05.md`
- `Operation_Clearwater_TwoDayMasterReviewRemediation_2026-06-28.md`
- `Operation_Causeway_CrossFsBridgeSecurityAndPerformanceRemediation_2026-07-07.md`
- `CodeReview_Last3Days_IndependentFindings_2026-07-11.md`
- `CodeReview_Last3Days_Remediation_2026-07-11.md`
- `UI_FileOperationsPopupCodeReviewRemediation_2026-07-10.md`
- `UI_FileOperationsPopupReviewFindings_2026-07-09.md`
- `Operation_Keystone_ThreeDayReviewRemediation_2026-06-20.md`
- `Operation_Farsight_ViewerPluginsRemediation_ExportSafetyDecodeReliabilityWebSecurityAndTextGeometry_2026-06-16.md`
- `Operation_ThreeDayReview_FolderView_CONTINUATION_2026-07-11.md`
- `Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md`
- `Operation_Firebreak_ConsolidatedReviewRemediation_2026-07-06.md`
- `Theme_ExpressivePaletteAndReferencesPlan_2026-07-13.md`
- `RedConfigure_LocalizationThemeManagerPlan.md`
