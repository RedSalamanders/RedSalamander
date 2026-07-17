# Specs/Plans/WIP - Index & Ranked Queues

Last WIP triage update: **2026-07-17** after Observatory and the independent last-four-days review closed, with
Astrolabe and Rosetta Lantern retained as the unique routed owners.

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

`CodeReview_Last4Days_IndependentFindingsAndRemediation_2026-07-16.md` is complete, merged to `master` in
`177600a8d`, and archived under `Specs/Plans/Done/`. Its session-end, queued-payload, theme-grammar/contrast, and
inline-theme preservation defects are closed. RedConfigure/gate/settings/Vite rows were reconciled to completed
Observatory tracks. Only the work it explicitly routed remains live: source-contract architecture in Astrolabe and
Czech/Japanese/Slovak linguistic review in Rosetta Lantern.

`Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15.md` is complete and archived under
`Specs/Plans/Done/`. It remains the 2026-07-15 whole-repository finding/disposition ledger, not a live queue.
Lighthouse is also complete and archived under `Specs/Plans/Done/`;
Track 0 is complete by bounded convergence: CI is green, Full's Track 0 paths and sandbox audit are green, six
unrelated Full failures cleared in the one focused batch, and the two reproducible unrelated cases are routed to
Tracks 9 and 6. Track 1 is complete: releases now require the exact requested artifact matrix, every action is pinned,
dependency/ownership guards are active, and six mismatched Squad workflows are removed. Observatory then closed
Track 2 with UIDVALIDITY-qualified identity, UIDPLUS-only single-message expunge, rollback, fake-mailbox
coverage, and archived Release perf evidence. Track 3 is also complete with typed Graph/upload URL boundaries,
redacted diagnostics, secure transient storage, continuation/redirect rejection, focused Debug/Release proof, and
archived request-count evidence. Track 4 is complete through the archived IronLedger owner with bounded drop data,
stable destination/provider identity, honest asynchronous MOVE reporting, focused runtime/source-contract proof,
and authoritative FolderView/File Operations contracts. Track 5 is complete with Pack overlap/cancellable deletion,
the Common local-file transaction, strict Monitor export, explicit Unpack conflict policy, and responsive Make File
List save/cancel/progress behavior. Track 6 code now has bounded queue/history/search, safe clipboard publication,
snapshot Document reads, owned layout-worker COM state, cancellable budgeted file open, complete Clear reset, and a
per-session ETW callback/handle ownership model. Track 6 is complete with deterministic shutdown proof, a green
same-machine latency archive, 139/139 source contracts, and authoritative Monitor/shared-helper contracts.
Track 7 is complete with conflict-aware unique-sibling settings publication, a real cross-process CAS proof,
save-blocked recovery when backup fails, strict numeric/Unicode serialization, explicit Preferences partial-success
behavior, a clean Debug rebuild, 140/140 source contracts, and authoritative SettingsStore/Preferences contracts.
Track 8 is complete: canonical connection IDs, collision-safe no-guess credential migration, purpose/secret-kind
scoped authorization, and lock/session clearing have focused build/runtime/schema/localization/source-contract proof.
Track 9 is complete with reentrancy-safe menu/tree dispatch, stable Tree/UIA identity, text-range conformance, and
focused DxUi proof. Track 10 is complete with shared bounded pagination, Graph upload acknowledgements, race-safe
Microsoft merge semantics, commit-aware cleanup, bounded Google transport/token refresh, reversible Google identity,
archived request/resource evidence, and a clean Debug x64 recovery rebuild. Track 11 is complete with exact S3
transfer proof, callback-result preservation, bounded recursive delete, commit-aware cleanup debt, retryable
multipart reconciliation, explicit AWS lifetime/unload state, and repeated Debug/ASan runtime-refresh proof.
Track 12 is complete with transactional configuration across all shipped providers, legacy secret scrubbing,
shared Curl/Google process-runtime ownership, and cancellable/generation-gated/bounded 7z indexing.
Track 13 is complete with fail-before-change local writer semantics, bounded path expansion, atomic Dummy
copy/move preflight, descendant read-only delete policy, and retained reversible Google identity proof.
Track 14 is complete with serialized one-shot final settings persistence, stale edit-suggestion rejection,
cached/viewport-bounded large-menu painting, capped sibling enumeration, focused Commands coverage, and archived
same-machine performance evidence.
Track 15 is complete with one declarative runtime-dependency manifest, fail-closed MSBuild/package validation,
fresh-extraction app/all-plugin smoke, separately pinned vcpkg tool identity, reconciled ARM64/critical PR gates,
and a scheduled/high-risk ASan lane with a seeded detector-health proof.
Track 16 is complete with composition-root ownership, independent production translation units, and compiled
RedConfigure page presenters with typed theme origin. Track 18 closed audit-governance drift and generated Vite
state; Track 19 closed typed/localized RedConfigure preview, validation, duplicate, locale, Undo, and performance
contracts. Track 17's independent source-contract disposition/selftest-boundary work is now owned only by
`Operation_Astrolabe_TestContractArchitectureMigration_2026-07-17.md`.

The unfinished CI auto-format fork/race hardening remains a separate live owner in
`../../../plans/012-format-autocommit-race.md`. Observatory deliberately coordinated but did not absorb that
P3 workflow-only task; use the legacy plan as its implementation detail until it is executed and closed.

## Best Next Pick

If the current goal is remediation or stabilization, take **I2 FIR-1 / FG-A1** next: add destination-side bridge
`GetSize` fault injection once and share that seam with Floodgate. The four-day delta review is complete; do not
reopen its closed payload/settings/theme rows. Astrolabe is lower-urgency test architecture and should run as a
dedicated slice. Rosetta Lantern can begin its exact inventory/report work, but translation completion depends on
competent target-language reviewers.
Do not reopen the closed Track 5 output, Track 6 Monitor work, or Track 7 settings coordination.
IronLedger, Lighthouse, Firebreak, Causeway, Farsight, and the
test-suite stabilization plan are completed references; do not reopen their routed rows.

If the current goal is product work instead of remediation, start with **N1 Google Drive Phase 2**:
interactive OAuth PKCE sign-in plus read-stream/download support.

## Improvement / Remediation Queue

Ranked by impact, urgency, dependency value, and whether the work is already owned elsewhere.

| Rank | File | Why this rank | Next action / routing |
|------|------|---------------|-----------------------|
| I2 | `Operation_FileOperations_FaultInjectionRatchet_2026-07-05.md` | Enables RED-able coverage for bridge and cloud data-safety gaps; useful before several FileOps fixes. | **FIR-1 / FG-A1**: destination-side bridge `GetSize` fault injection through the bridge IO decorator. |
| I3 | `Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md` | Cloud parity/perf proof queue; overlaps the coverage seam needed by FIR. | **FG-A1**: destination-side `GetSize` fault-injection hook. Coordinate with I2 so only one seam is built. |
| I5 | `Operation_RosettaLantern_RedConfigureSatelliteLocalizationReview_2026-07-17.md` | Unique owner for F4-LOC-01; resource plumbing is complete, but ordinary RedConfigure UI text remains English in Czech, Japanese, and Slovak satellites. | Refresh the exact source-identical inventory and neutral-token allowlist, then obtain target-language translation/review before claiming completion. |
| I6 | `Operation_PerfMeasurementContract_2026-07-06.md` | Raises the quality bar for every perf-sensitive plan; not a runtime bug, but it prevents weak closeouts. | **Task 1**: install the Perf Measurement Record template in authoritative specs and add representative records. |
| I7 | `UI_FileOperationsPopupUxRefinementPlan_2026-07-07.md` | User-requested progress-popup clarity work; A-tier, queued-task Start now, footer bulk pause/resume, B4 completed grouping, and both review-remediation passes are implemented, but workflow/accessibility polish keeps the plan in WIP. | Continue remaining backlog: B1 reorder, B4 completed-group animation, B5 keyboard/UIA, C5 Keep Both engine support, C7 shared DxUi graph closeout. |
| I8 | `Operation_Astrolabe_TestContractArchitectureMigration_2026-07-17.md` | Independently owns reviewed source-contract dispositions, behavioral companions, residual suite splitting, and optional selftest compilation boundaries; not a release blocker. | Replace the regex-shape heuristic with one explicit reviewed disposition per live source-contract case before deleting or splitting anything. |
| I8A | `../../../plans/012-format-autocommit-race.md` | Separate P3 CI workflow owner retained from the historical root-plan set; Observatory Track 1 explicitly did not absorb it. | Add fork/event, stale-head, and concurrency guards, then verify the workflow on GitHub before closing the legacy plan. |
| I9 | `FileSystem_CrossFsBridgeImplementationAndPerformanceRedesign_2026-07-07.md` | Design-level performance rethink for the cross-FS bridge; valuable after the immediate Causeway fixes. | Use as a dedicated design/perf slice; start with R-1 depth-N pump only with perf evidence. |
| I10 | `Win32_Inventory.md` | Narrow RAII/lifetime/DPI audit with bounded fixes. | **CHK-1**: replace `DestroyWindow(_hWnd.get())` on `wil::unique_hwnd` owners with reset/transfer-safe close. |
| I11 | `Operation_Tailwind_FolderViewWarpDrivePreMergeReviewRemediation_2026-06-30.md` | Residual branch self-review ledger; many items are done or routed, so it is not the best primary queue. | Reconcile freshness first; execute only remaining non-routed items after WarpDrive/DxUi owners confirm no overlap. |
| I12 | `Core_HResultStatusFormattingCleanup_2026-06-19.md` | Low-risk cleanup and localization consistency, but low urgency. | **WI-1**: inline `FormatHResult` and `FormatStatusText`; coordinate with hot FolderView files. |
| HOLD-I13 | `Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md` | Active/coordination-sensitive closeout plan. | Do not pick unless the WarpDrive owner hands it off. |
| HOLD-I14 | `FolderView_WarpDrive_ContinuationBaton_2026-06-29.md` | Resume state for WarpDrive closeout and suite proof. | Do not pick unless the WarpDrive owner hands it off. |
| HOLD-I15 | `DxUi_Uia_ContinuationBaton_2026-06-29.md` | Active/coordination-sensitive UIA snapshot and commit-split baton. | Coordinate before touching DxUi/UIA snapshot coalescing or related routed Tailwind/Granite items. |
| HOLD-I16 | `Notes_Scratch.md` | Human-managed scratch file, not an executable agent queue. | Keep out; owner can delete or reconcile manually. |

## Deferred Decision Review Queue

These operations preserve previously rejected or policy-sensitive scenarios for an explicit maintainer decision.
They are not implementation queues; approved work must be routed to a separate executable WIP owner.

| Rank | File | Why it is retained | Next action / routing |
| --- | --- | --- | --- |
| D1 | `Operation_Parallax_DeferredArchitectureSecurityAndSimplificationDecisionReview_2026-07-14.md` | Preserves twelve Lighthouse rejection/policy scenarios with explanations, evidence requirements, reopen triggers, and decision options. | Review PAR-1 through PAR-12 when the relevant security, product, or architecture policy is ready for decision. Record one outcome per scenario and route approved implementation separately. |

## New Feature Queue

Ranked by concrete implementation readiness and product value. These are user-visible additions or
feature-roadmap plans, not primarily bug-fix/remediation queues.

| Rank | File | Why this rank | Next action / routing |
| --- | --- | --- | --- |
| N1 | `FileSystem_GoogleDrivePluginPlan.md` | Concrete new filesystem plugin; enumeration landed, but no in-product first sign-in or file reads yet. | **Phase 2**: interactive OAuth PKCE sign-in plus read-stream/download support, using Microsoft Drive as the local model where applicable. |
| N2 | `FolderView_ThumbnailBackgroundEnrichmentFollowup_2026-07-04.md` | User-visible completion of the cached-only thumbnail decision; restores richer cold-cache thumbnails without blocking first paint. | Implement bounded background enrichment for local-shell-backed files with generation checks, metrics, and perf evidence. |
| N3 | `Terminal_EmbeddedConPtyLibGhosttyVt_IntegrationPlan_2026-07-13.md` | Implementation-ready embedded terminal architecture with explicit ConPTY, VT, lifecycle, security, accessibility, and perf gates. Large scope makes it less suitable than N1/N2 for the immediate next slice. | Start only as a dedicated branch from its Phase 0 capability ledger and dependency/license gates; do not mix it with FolderView or DxUi hold work. |
| N4 | `Product_WhimFilesGapAnalysisAndImprovementPlan_2026-07-08.md` | Product roadmap/gap analysis, not an implementation-ready slice. | Convert the highest-value gaps into separate executable plans before coding. Start with G1 live multi-dimensional filtering or G2 undo if product direction is approved. |

## Archived Reference Ledgers

These were consolidated or completed and should stay in `Specs/Plans/Done/` unless fresh evidence
reopens a specific item:

- `Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15.md`
- `CodeReview_Last4Days_IndependentFindingsAndRemediation_2026-07-16.md`
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
- `Operation_CommandsSelfTestInputIsolation_2026-06-24.md`
