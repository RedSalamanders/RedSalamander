# I29 File Operations post-merge qualification

Status: ACTIVE — qualification and evidence closeout only.
Owner: Codex, this continuation. Activated by the user's 2026-10-05 approval to
archive delivered implementation separately from unfinished qualification.

## Scope and authority

The [I27 safety implementation](../Done/UI_FileOperationsSafetyRemediation_2026-10-01.md)
(the archived File Operations safety plan; I27 is never reused) is delivered and merged. This plan receives its three unfinished closeout gates;
none is waived or treated as passed. Do not resume the completed implementation
plan, reopen delivered fixes without a demonstrated failure, or launch overlapping
Full runs simply because more than one plan needs their results.

Current contracts remain [File Operations](../../FileSystem/FileSystem_FileOperations.md),
[popup](../../UI/UI_FileOperationsPopup.md), [FolderView](../../UI/UI_FolderView.md),
[shared grid/alerts](../../UI/UI_DxUiSharedGrid.md),
[shared helpers](../../Core/Core_SharedHelpers.md),
[selftests](../../Testing/Testing_SelfTests.md),
[coverage](../../Testing/Testing_TestCoverage.md),
[performance validation](../../Testing/Testing_PerformanceValidation.md),
[validation evidence](../../Testing/Testing_ValidationEvidence.md) and
[spec authority](../../README.md). Domain specs stay in their canonical folders.

Independent owners remain independent: [DxUi I31](DxUi_HostMessagesToolsetAndRestore_2026-10-01.md),
[confirmation/outcomes I18](UI_ConfirmationPreferencesAndOperationOutcomes_2026-09-08.md),
[stress validation I3](FileOperations_RemainingStressValidation_2026-08-04.md),
[provider/harness residuals I25](FileOperations_S3PathStyleCopyAndHarnessResiduals_2026-09-15.md)
and [deferred platforms H4](DxUi_DeferredPlatformQualification_2026-09-13.md).
The closed [release recovery I28](../Done/Build_ReleaseDependencyRecovery_2026-10-02.md) no longer owns Full qualification; its unfinished October 2 cases, including Compare selection, the stale conflict prompt, the capability ceiling, and the ViewerPE Close timing, are owned here together with I31's consumer Full.
This split closes none of the other owners' work. Native ARM64/ASan, library qualification,
current-pin CI and release handoff remain with their named owners.

## Published starting point and exact evidence boundary

| Item | Verified delivery state |
|---|---|
| Published `master` | `cfe94556c94e971a753d692c55793eddda172bd7` |
| Delivered implementation | `02841dd4b62022a5c3d04f72efa18a3dea84a2bb` |
| Reviewed integration | `6e54d8679fa110bb2b8364ac306404788e641638` |
| Dependency | [DxUi lock](../../../Dependencies/DxUi.lock.json), API revision 3 at delivery; inspect the actual current lock before testing |
| Integrated Debug build | Whole solution, x64, test-enabled Embedded; zero warnings/errors |
| Build receipt | `0bfec155d5d67fae48227cd791b1edbca1e7317c5e9e2887e3f63a4059f01f17` |
| Compiled workspace source | `565bac37210db3e3c0df6b66b5867d7dc831527d28a4302ad6e1743cf2ed7a01` |
| Two integrated Debug cases | 2 passed, 0 failed, 0 skipped; repeat and timeout scale one |
| Native run | `20261005T155340Z-46636-1f9c0561c42245cbb9d571b1ad07655d` |
| Reviewed original archive | Six byte bindings; manifest `5ce1b427218e3a70a66b9d5e06fece801c145f98c7d544b4ebea19cab6f2350e` |
| Focused merge tooling | 255 passed, 0 failed, 0 skipped; no blocking spec/tool inventory findings |

The two native cases are
`cmd_pane_fileops_popup_throughput_disclosure` and
`cmd_pane_fileops_popup_failure_surface_on_d2d_target_failure`. They establish
exactly those integrated Debug scenarios. They do not establish Full, repeated
reliability, current-pin resource equivalence, or the cause of prior failures.
The later checkpoint/index and this delivery split change no compiled inputs.

Latest complete Debug and Release Full runs remain **FAILED**. Debug included
the earlier disclosure/failure-surface fixtures, an F7 foreground refusal and
ViewerPE input/restoration. Release included a settings-child timeout, theme
timing, R0e timing and ViewerPE restoration. Later focused passes do not rewrite
those verdicts. Original run identities, immutable receipt epochs and closed
archives remain in the delivered plan and external retention root below. The
historical 265-entry disk audit is not a clean audit; 65 informational spec
findings are distinct from blocking inventory findings.

## Remaining work and acceptance

- [ ] **Q1 — Integrated Full outcomes.** Inspect current `master` ancestry and
  coordinate the required x64 Debug/Release Full runs with the DxUi/consumer
  owner. Reuse an existing actual same-source/profile run when it satisfies the
  required contract; one run can supply evidence to multiple plans. Retain
  canonical builds, build receipts, complete per-case verdicts, unchanged
  limits, source/pin/module identity and reviewed no-follow archive closure.
  Preserve skipped prerequisites and all failed originals. Resolve demonstrated
  failures on current code with focused checks first; broaden or repeat only
  when a changed binary, failure or uncovered acceptance requirement warrants
  it. A selected case, composite continuation or green build is not a fresh Full
  pass. Full remains required by the current contracts, but was deliberately
  moved after delivery rather than blocking the implementation merge.
- [ ] **Q2 — Performance and durable evidence.** Publish the already-reviewed
  final performance packages under `Specs/TestRuns` after verifying their exact
  retained bindings and archive limits. Preserve API-2 provenance and bounded
  acceptance: these older packages do not qualify API-3 adoption. Determine
  which current-pin paired Release evidence is still needed with the consumer
  owner, using unchanged fixtures and retained baseline/candidate processes.
  Preserve adverse directions, unequal dependency identities, lifetime/driver
  limits and capability skips. Do not invent neutrality or silently rebaseline.
  Keep the accepted F6/v2/F7 scope below; seek developer advice only for a new
  confirmed degradation not covered by an existing explicit decision.
  - [ ] **Q2-L — Local route-observation qualification.** The Local provider's
    guarded following-handle route observation (`TryObserveCachedLocalRoute`,
    which reuses the cached volume route after a fresh endpoint and mount GUID
    agreement) is delivered with synthetic Debug controls only. Its native
    correctness and paired timing evidence is owned here: mounted-volume,
    junction-retarget and remount behavior on the actual Local route, plus
    unchanged-fixture paired timing against the original volume resolver, with
    source-bound receipts. The mounted route stays unqualified while Windows
    denies mount registration; do not infer it from synthetic controls. If the
    privilege stays unavailable, record the gap as an explicit owner decision
    rather than leaving it implicit. The File Operations contract and developer
    guide route this obligation here.
- [ ] **Q3 — Final contract and governance closeout.** Reconcile affected domain
  specs, shared-helper catalog, coverage/evidence and real consistency-ledger
  anchors against the integrated implementation and reviewed outcomes. Pass
  spec/tool inventories, SpecInformationArchitecture,
  DocumentationDriftContracts and relevant formatting/whitespace checks. Reuse
  Q1's actual Full results for the specification-governance Full requirement.
  Validate changed archives and the contract-era archive inventory; retain
  every failure. Admit this successor's exact Done path and a focused test that
  rejects an unreviewed sibling, then move this plan to Done and retire only I29
  from the WIP index. Leave existing Done history and other owners untouched.

The accepted [F6/v2 assessment](../../TestRuns/4cb089111a23/FileOps/20261003T211512Z-80796-0cb73e59c3c741f9bcae95a38bb43929/review/ROOT-F6-V2-BOUNDED-ASSESSMENT.md)
and [F7 assessment](../../TestRuns/4cb089111a23/FileOps/20261003T191317Z-82816-199e9f25af4d49739d1410c6881e3e0d/review/ROOT-F7-ASSESSMENT.md)
retain actual timer/active-work observations, descriptive resource differences,
252 semantic admission pairs, 126 prime-ingress pairs and their adverse samples.
All 24 actual mounted samples were capability skips after Windows denied mount
registration. That route remains unqualified; synthetic controls do not replace
it. No new privilege or provider-policy change is authorized by this plan.

## Retention and execution checkpoint — 2026-10-05

Delivery publication and independently reviewed integrated receipts are retained
under the authorized external root:

`C:\RedSalamander.Perf\evidence\i27-followup-20261002-0e139fb\resume-20261003T121349Z\final-run-review\`

Within `resume-20261005T060429Z/`:

- `merge-convergence-20261005T153458Z-a1e68395/` contains publication,
  independent integration review, compiled receipt and focused test originals.
  The retained build receipt is 83,539 bytes, SHA-256
  `e31d41c5e10747e993e3fa199c6415e4889a5ff11488082e70aba2f06e942185`.
- `final-perf-publication-prepared-20261005T104314Z-faf156e65f84495c81094248bc33210f/`
  contains 22 prepared packages, 348 files and 31,472,181 bytes, not yet final
  publication. `final-perf-publication-root-review-20261005T110049Z.json`
  retains root's 325 literal pair-binding checks. Keep the API-2 historical
  scope and originals; do not restamp these as current-pin results.

Current implementation checkout: `C:\RsI27Followup`, branch
`codex/fileops-safety-followup`; canonical master: `Z:\src\RedSalamander`.
Both published heads were verified at `cfe94556c` before this documentation-only
split. No Full/native/build session is launched by this checkpoint.

Delivery-split validation: SpecInformationArchitecture 37/0/0 and
DocumentationDriftContracts 18/0/0; all 77 local links across the three changed
documents resolve. Spec/tool inventories have zero blocking findings, with 65
informational spec findings retained and protected Done history valid. Root
checked the complete change scope, links and retained receipt hash independently
of the cheaper read-only review. Logs and structured results are retained under
`resume-20261005T060429Z/delivery-split-20261005T164845Z-d9939707/` within the root
above. This is metadata validation, not completion of Q1–Q3.

**Next action:** check current master and the existing consumer owner's actual
qualification results; map usable same-source/profile evidence to Q1/Q2 before
starting any new broad run. Keep a single writer/build owner. Use a cheaper
agent for bounded read-only evidence or documentation review when useful;
independently verify its claims. Protected native input uses the existing
warning, interactive desktop lease and focus/cursor restoration at an agreed
available time. Application captures use only deterministic harness scenarios.

The implementation is delivered. This successor remains ACTIVE until Q1–Q3
are satisfied with their actual outcomes and permitted prerequisite boundaries.
