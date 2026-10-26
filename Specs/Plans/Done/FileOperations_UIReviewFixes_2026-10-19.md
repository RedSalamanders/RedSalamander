# File Operations UI review fixes

Scope: close all four findings from the 19 October review while preserving the accepted card hierarchy and Last note wrapping.

## Checklist

- [x] Replace every obsolete combined-summary expectation in Commands throughput tests with independent speed/ETA region assertions and exact localized French discovery facts; preserve placement, graph count, persistence and idempotence checks.
- [x] Bind live count labels to independently expected totals and numeric facts from the actual paint generation; verify retained accessibility and reject incorrect labels.
- [x] Require scenario-specific percentage presence and exact localized value for determinate cards, with explicit absence for open/unknown/skipped totals.
- [x] Restrict the automatic transfer-concurrency note to Copy/Move and cover Delete/manual/unknown/finished absence.
- [x] Synchronize the normative spec, user guide and test contract.
- [x] Repair the focused source tooltip cache found by foreground validation: other-stream text changes refresh it even when path and bounds stay fixed.
- [x] Build, run the relevant no-activation and foreground Commands tests, inspect native captures and archive evidence.
- [x] Review the final result again, record limits and move this plan to Done.

Protected rendering scenario: the existing English/French 480/960 DIP card matrix, including expanded Delete. Reuse FileOps.Popup.CardHierarchy.FrameUs and the popup's existing measurement instrumentation. Use the exact pre-fix Last note candidate as same-machine baseline where scene behavior matches; identify the intentionally changed Delete note separately. No storage-throughput or performance-equivalence claim.

Validation: no-activation FileOps_ReviewRemediationContracts and DiscoveryScope_FolderLiveProgress; governed foreground Commands throughput disclosure, external reload and restart cases. The user explicitly agreed to the foreground tests after the build finishes during this task. Use the runner's existing warning and desktop lease. Preserve the unrelated unstaged DxUi pin. Keep source files stable until each complete runner exits.

## Closeout

Final x64 Debug solution build passed with zero warnings/errors. Receipt:
`ceb646bc0fd1e47dc2c9cf61cf1ea4110c3e4281814923330281745ef84b866e`.
The final candidate passed all three Commands throughput cases, both FileOps scenarios (six
runner cases), 69 live-operation assertions and 204 source contracts. The matrix produced 132
native captures and passed 124 explicit card-fact checks. All final runtime runs used this same
unchanged candidate; only plan closeout and evidence/staging followed.

Runtime qualification also repaired obsolete French discovery wording, an Escape/refocus fixture
assumption, idle-popup setup and a real focused-source tooltip cache defect when another stream
finishes without changing the displayed path or bounds. Original failures remain archived.

Matched warm rendering p95 was 8.550 ms before and 8.575 ms after; matched path-access batch p95
was 66.852 ms before and 66.746 ms after. Single observations do not prove equivalence or causality.
The effective-concurrency change intentionally removes the Delete note; the additional unknown-limit
scene is separate from matched rendering figures.

The [normative UI contract](../../UI/UI_FileOperationsPopup.md) and user guide own the durable
behavior. [Evidence and final review](../../TestRuns/4cb089111a23/FileOps/2026-10-19_ui-review-fixes/README.md)
bind runs, sources, captures, metrics and failed attempts. No remaining actionable defect was found
in the final scoped review. Full, Release, ARM64, ASan, other DPI scales and external assistive
technology remain unevaluated. The shared sandbox audit's 274 findings remain unresolved; selected
case passes are not clean whole-repository qualification. The unrelated DxUi pin was untouched.
