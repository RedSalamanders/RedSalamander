# I58 Planning lifecycle and owner reconciliation

> **NON-NORMATIVE ACTIVE PLAN.** Single documentation and governance remainder of the retired August 25 advisor survey. Product behavior stays with current domain contracts and indexed product owners. The [retired survey index](../Done/AdvisorSurvey_Index_2026-08-25.md) preserves the five original plans and their provenance.

## Status and scope

- **State:** ACTIVE; consolidation delivered, remaining documentation work below is unimplemented.
- **Priority:** P2 documentation and planning governance.
- **Owner:** repository specification maintainers; coordinate edits with each affected indexed owner.
- **Planned at:** `269b06fc9f79f9fdd2bcbe1fce6423ed3b1424f1`, reconciled 2026-11-03.
- **Ownership boundary:** lifecycle metadata and its bounded inventory guard, contradictory current-plan directions, and current owner routing. Existing owners retain implementation, human review, qualification, budgets and acceptance decisions.
- **Authoritative contracts:** [specification authority and plan lifecycle](../../README.md), [tooling governance](../../Testing/Testing_ToolingGovernance.md), [agent guidance](../../Testing/Testing_AgentGuidance.md), [validation evidence](../../Testing/Testing_ValidationEvidence.md), and [performance validation](../../Testing/Testing_PerformanceValidation.md).
- **Implementation/tests:** extend the existing `Tools/Modules/Tooling/SpecInformationArchitecture.psm1` and `Tools/Tests/SpecInformationArchitecture.Tests.ps1` for the reviewed lifecycle-field guard. The bounded historical-review metadata defect in `Tools/Tests/DocumentationDriftContracts.Tests.ps1` is a validation prerequisite below. Current archive-path admission is already delivered by this consolidation.
- **Drift check:** run the following and inspect dirty/untracked state before continuing. Changed excerpts require fresh evidence; the August survey's line numbers and proposed edits are historical.

```powershell
git status --short
git diff 269b06fc9f79f9fdd2bcbe1fce6423ed3b1424f1..HEAD -- Specs/Plans/WIP/README.md Specs/Plans/WIP/PlanningLifecycleAndOwnerReconciliation_2026-11-03.md Specs/Plans/WIP/FileSystem_GoogleDrivePluginPlan.md Specs/Plans/WIP/FolderView_ThumbnailBackgroundEnrichmentFollowup_2026-07-04.md Specs/Plans/WIP/Operation_Tailwind_FolderViewWarpDrivePreMergeReviewRemediation_2026-06-30.md Specs/Plans/WIP/Operation_PerfMeasurementContract_2026-07-06.md Specs/Plans/WIP/Operation_Astrolabe_TestContractArchitectureMigration_2026-07-17.md Specs/Plans/WIP/Operation_RosettaLantern_RedConfigureSatelliteLocalizationReview_2026-07-17.md Specs/Plans/WIP/Core_HResultStatusFormattingCleanup_2026-06-19.md Specs/Plans/WIP/Win32_Inventory.md Specs/Plans/WIP/Product_WhimFilesRemainingDecisions_2026-08-25.md Specs/Plans/WIP/Operation_Atlas_RemainingSpecificationDecisions_2026-08-04.md Specs/Plans/WIP/Operation_Parallax_DeferredArchitectureSecurityAndSimplificationDecisionReview_2026-07-14.md Tools/Modules/Tooling/SpecInformationArchitecture.psm1 Tools/Tests/SpecInformationArchitecture.Tests.ps1 Tools/Tests/DocumentationDriftContracts.Tests.ps1 Specs/Reviews/WIP_FolderQuality_2026-11-01.md Specs/README.md
git diff --stat
```

## Reconciliation and ownership

The old survey was generated on 2026-08-25 at `b5c3b0eab`. Its separate root directory followed contradictory former Improve instructions. The October 28 skill repair requires indexed WIP output. This consolidation retires all five old TODO rows; it does not declare their product subjects finished.

| Original survey | Verified current state | Disposition and remaining owner |
|---|---|---|
| [001: lifecycle/status](../Done/AdvisorSurvey_001-wip-lifecycle-and-status-truth_2026-08-25.md) | Google Drive, thumbnail enrichment, Astrolabe, perf measurement, Rosetta and HRESULT have incomplete lifecycle metadata. The inventory enforces index/rank/authority rules but has no required lifecycle-field guard. | R1 owns documentation normalization and the bounded guard. Keep existing product IDs and authority. |
| [002: FolderView routing](../Done/AdvisorSurvey_002-warpdrive-tailwind-executable-surface_2026-08-25.md) | I2 already has the compact remainder, Task 10 verification-or-waiver, Task 11 and separate I1 ownership. H0 has a correct HOLD banner but its body still sends TW-5/23/24 to retired batons and sequences the already-present paste-shortcut generation check. | Retire the master/split/banner prescriptions. R2 owns contradictory body routing; I1/I2 retain feature/performance work, and H0 remains evidence-only. |
| [003: closeout gates](../Done/AdvisorSurvey_003-closeout-gate-hygiene_2026-08-25.md) | I12 completed on September 8 and is in Done. I10 already distinguishes delivered implementation from its unmatched Full evidence. I5 still instructs editing the former WIP flake-convergence plan and carries older template/test-root instructions. | Retire I12 instructions and duplicate I10 status work. R3 reconciles I5's handoff. I2/I10 and other qualification owners keep their evidence gates; no waiver is granted. |
| [004: active-owner refresh](../Done/AdvisorSurvey_004-refresh-stale-active-owners_2026-08-25.md) | Google Drive implements reads/writes/mutations; its WIP still calls them stubs. The FolderView HRESULT alias exists, while the old Issues-pane alias name is absent. Win32 was refreshed October 17 and retains test-only CHK-1 sites. I9's implementation is superseded by I55, with residual acceptance in I9/I56. Rosetta retains a historical English-only premise requiring a fresh inventory. | R4 refreshes stale owner descriptions. Retire August Win32/CI implementation recipes. I0/I3/I4/I6/I7/I8/I9 retain their respective product, review or validation work. |
| [005: decisions/HOLD/scratch](../Done/AdvisorSurvey_005-decision-hold-scratch-routing_2026-08-25.md) | D1 already forbids treating its waves as a schedule, and D0 records accepted clipboard/Batch Rename/ABI decisions. D1 G2 still points at retired I11. H1 retains PENDING REVIEW dispositions. | R5 refreshes decision routing with evidence. Retire duplicate decision banners and the scratch-derived routing task. The exact human-managed note is excluded and was not inspected during consolidation. |

Supporting current evidence:

- [Google Drive contract](../../FileSystem/FileSystem_GoogleDrive.md) records the implemented writable-provider milestone; interactive sign-in and remote qualification remain distinct.
- `RedSalamander/FolderView.FileOps.cpp`, `FolderView::OnPasteShortcutComplete`, computes `generationMatches` and uses it when applying completion. Do not reimplement the old TW-7 recipe.
- [I12 completed record](../Done/Operation_ReviewFollowup_FOTerminalFocus_2026-08-19.md) retains its September 8 closeout and failures. It is frozen history.
- [I10](Tooling_ToolsInventoryAndGovernance_2026-08-13.md) records the October 17 source-identity check. An unrelated aggregate Full cannot close it.
- [I9](CI_FormatAutocommitRace_2026-08-25.md) routes shared formatting implementation to [completed I55](../Done/CI_OrganizationStandardization_2026-10-29.md) and hosted acceptance to [I56](CI_HostedAcceptance_2026-10-31.md).
- [Win32 inventory](Win32_Inventory.md) already documents external DxUi ownership and the remaining test-only owned-HWND calls. Keep its current evidence instead of restoring August claims.

## Ordered remaining work

### R1: normalize lifecycle metadata and guard the agreed fields

- [ ] Reconcile status, planned-at identity, drift check, ownership, remaining outcomes, normative references, implementation/tests, verification, done criteria and STOP conditions in the six legacy ACTIVE owners named in row 001, plus other indexed plans demonstrably missing current required fields. Exclude the exact human-managed note before opening content.
- [ ] Preserve valid existing metadata, priorities, stable IDs, HOLD/DECISION boundaries and named product owners. Historical progress such as “WIP” may remain narrative, but the machine state must agree with the index.
- [ ] Extend the canonical specification inventory with the smallest reviewed lifecycle check that enforces the current contract. Accept supported table/list forms; do not require a new template or second parser. Document any precise legacy exception beside its owner and test it.
- [ ] Add valid/missing-field/index-state controls in the existing IA suite. Keep malformed-state detection separate from semantic freshness and runtime qualification.
- [ ] Repair the historical-review provenance guard described in the validation record below. Preserve the common frozen-history/index-routing requirements and each reviewed campaign's actual date/commit/owner; restrict July-specific assertions to their reviewed cohort. Reject missing or incorrect provenance with focused controls. Do not edit frozen reviews to insert another campaign's metadata.

### R2: remove contradictory Tailwind execution directions

- [ ] Re-read H0 against the actual restored DxUi identity and current FolderView source. Reconcile TW-5/23/24 baton assignments, the old TW-7 sequence, and stale in-tree `Common/DxUi` paths with the HOLD banner.
- [ ] Keep I1's thumbnail enrichment and I2's remaining perf closeout separate. Route a freshly confirmed uncovered repair to an existing accepting owner, or leave it HOLD pending explicit promotion. A stale finding alone does not establish a current defect.
- [ ] Preserve rejected findings and useful evidence without executable directions to resume retired batons. Do not rewrite Done history.

### R3: rebase the performance-measurement handoff

- [ ] Reconcile I5's template and field-name prescriptions with the current normative performance contract. Remove instructions to edit flake-convergence as a WIP; link its Done record for historical evidence only.
- [ ] Replace obsolete sandbox/setup directions with the current owned-sandbox contract. Requested environment paths do not authorize initialization or another volume.
- [ ] Classify each remaining enforcement item as delivered, missing, or owned elsewhere using the actual runner/analyzer/tests. Keep genuine enforcement work in I5; do not implement it in this documentation slice.
- [ ] Keep I2's perf matrix and I10's matching Full obligations with those owners. Preserve failed outcomes and acceptance boundaries; no duplicate Full, budget substitution or silent waiver.

### R4: refresh stale active-owner facts

- [ ] Rebase I0's read-only/stub descriptions and phase boxes on the Google Drive contract, capabilities and executing provider-contract tests. Keep initial sign-in, remote proof and product choices explicit; delivered mutations are not new work.
- [ ] Recheck I7's aliases/callers and remove obsolete line-based claims. Existing UI/error work remains I7 and coordinates with I18.
- [ ] Rebase I6's English-only premise on a current resource comparison when that owner performs its inventory. Record existing translation versus remaining human review separately; resource parity never proves linguistic approval.
- [ ] Preserve I8's October refresh, I3's outcome-level stress ownership and I4's reviewed test-migration boundary. Refresh only drifted evidence; do not recreate their queues or execute old CI recipes.

### R5: reconcile decision and HOLD routing

- [ ] Update D1's retired I11 pointers using domain authority and accepting live owners. If implementation remains deferred with no accepting owner, state that in D1. Do not invent a new ACTIVE feature.
- [ ] Preserve D0's accepted clipboard, Batch Rename and ABI decisions. Refresh an unresolved row only where current source or ownership changed.
- [ ] Keep H1's unreviewed rows pending unless fresh scenario evidence supports an allowed disposition. Remaining test-only CHK-1 calls prevent treating all HWND cleanup as complete.
- [ ] Keep the exact human-managed scratch note excluded. Inherited extracts in the retired survey grant no read, edit, summary, routing or deletion authority. The current authority catalog requires a separate direct user instruction naming that file.

## Verification and acceptance

Use the current scoped workflow and an existing marked test sandbox. Run `Test-Changes.ps1 -Explain` first, then affected checks appropriate to the changed files. Use the governed runner for generated test data and saved results; inventory reports retain their documented repository-local `.build` output contract.

```powershell
.\Test-Changes.ps1 -Explain
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
.\Tools\Get-ToolInventory.ps1 -Validate
Invoke-Pester .\Tools\Tests\SpecInformationArchitecture.Tests.ps1 -PassThru
Invoke-Pester .\Tools\Tests\DocumentationDriftContracts.Tests.ps1 -PassThru
.\Tools\Test-Changelog.ps1
git diff --check
```

Expected: one index row for this plan, resolving current links, unchanged protected history, exact archive admission with unreviewed sibling rejection, consistent machine states, and focused tests passing. Structural checks do not establish semantic freshness, language approval, native/platform/foreground coverage, Full acceptance or performance acceptance. Product qualification remains with named owners and their contracts.

### Executed consolidation validation

On 2026-11-03, the canonical in-memory specification inventory reported 499 files and zero blocking findings. Protected Done validation reported zero modified, missing or unexpected files and exactly six admitted survey records. All six archived bodies matched their original SHA-256 values after reversing index navigation relocation. The human-managed note and existing Done history were unchanged. Tools inventory, changelog validation, PowerShell parsing, whitespace checks and formatter WhatIf passed; both formatter input sets contained the same 654 files.

The focused IA, documentation-drift, tool-inventory, tooling-governance and formatter run recorded **149 cases: 148 passed, 1 failed, 0 skipped**. The new six-path admission/sibling/companion/premature-closeout control passed. Preserve the failed result: `Historical audit governance contracts / marks every review campaign document as a non-live historical snapshot` fails at `Tools/Tests/DocumentationDriftContracts.Tests.ps1:300`. Its all-review loop requires `2026-07-17`, `f4e0c8c3bed8` and the Observatory owner from every review. The November 1 review has its own historical banner and `eb779e0bd836` snapshot, without that July date. Both the test and `Specs/Reviews/WIP_FolderQuality_2026-11-01.md` are unchanged from baseline `269b06fc9f79f9fdd2bcbe1fce6423ed3b1424f1`; the mismatch predates this consolidation. R1 owns the bounded guard repair; the focused run is not reported green.

Original XML, summary and preflight setup history remain under the existing marked sandbox at `D:\RedSalamander.Perf\runs\20261103T145541Z-44044-4c5513432bcb47dba473e35208eb689d\scratch\tools-pester\advisor-survey-consolidation`. `Test-Changes.ps1 -Explain` saved a broad advisory plan and returned its documented `NOT_EVALUATED` status. No native Full or another owner's qualification was executed or accepted by this consolidation.

### Consolidation delivered on 2026-11-03

- [x] Rechecked the five survey topics against current owners and targeted source.
- [x] Created this single indexed remainder and retired the five parallel TODO rows.
- [x] Preserved all six survey files in Done with source/blob/SHA-256 identities and recovery commands; only original index navigation was retargeted.
- [x] Removed the former root directory and updated live navigation, formatter exclusions and the current review's relocation note.
- [x] Added six exact archive paths to reviewed Done admission with sibling/companion rejection coverage. Execution results are recorded with this change's validation, not inferred from these checkboxes.

## Done and stop criteria

Done requires R1–R5 to be completed or explicitly routed with an accepting owner, no conflicting executable directions in the touched current plans, focused governance/affected checks passing, and durable rules merged into the authority catalog/tooling contract. Then move this plan to Done with exact admission and a rejection test. This plan does not stay ACTIVE waiting for another owner's product qualification.

Stop and preserve evidence if routing changes product policy, a delivery claim lacks matching evidence, another owner has conflicting edits, existing Done history would need rewriting, or a test needs a new root, desktop lease or environment permission outside scope. Record the issue under the existing owner; do not reopen the retired survey or create another root queue.

Recovery is a reviewable revert of this consolidation's path moves, metadata and exact-path admission together. Recover original survey bytes using the source commit/path in each archive; never overwrite concurrent work or treat recovered executor prose as current authority.
