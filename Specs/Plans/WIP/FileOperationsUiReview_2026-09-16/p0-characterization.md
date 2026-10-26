# P0 source characterization — 2026-09-19

RedSalamander characterization started at `e61e349c3e8d133e9d52a4c63ba4c49266bc4db5`. Drift from the planning base
`1425d56799bbc5fce158e92be93da92ef9982292` is documentation only; the popup, selftests and dependency
pin are unchanged. The consumer still pins DxUi `f5c7fc7403e352425b8e82b7a4da1c042d942ede`;
the canonical library checkout is `b125edbf4cdd639d1cabd686f9dc3fe8aab1f374`.

This is source characterization, not proof that every prior screenshot defect reproduces on the
current binary. Fresh runtime receipts and unresolved witnesses belong below as they complete.

## Capability and ownership decisions

| Need | Existing mechanism / finding | Next implementation owner |
|---|---|---|
| Ordinary task actions | Popup derives action height solely from conflict row heights; no conflict gives zero height. Snapshot flags currently accept a button record without checking its area. | Consumer: independent ordinary action height; strengthen the real popup smoke fixture's bounds/hit assertions. No library change is needed for this defect. |
| Text measurement | Public `DxUi::Typography` already has `MeasureSingleLineTextMetrics`, `MeasureSingleLineTextWidthPx` and `MeasureWrappedTextHeightPx`; Label supports multiline. | Reuse these services and current host typography; do not add a second measurement factory/cache. Qualify font/DPI consistency and failure handling. |
| Measured action flow | StackPanel accepts caller-specified extents and does not wrap. Button paint passes single-line text to DrawCenteredText; no public multiline Button option exists. | DxUi: shared action-flow geometry and opt-in measured multiline Button support, preserving old defaults. Consumer supplies labels, order, eligibility and minimums. |
| Repeat-scope checkbox | DxUi has Checkbox and Toggle semantics. Popup creates ordinary Buttons for every action and paints the old checkbox beneath hosted controls. | Consumer: host a real Checkbox, synchronize checked state, keep stable action identity and avoid duplicate callbacks. Qualify library checkbox behavior before changing it. |
| Disclosure accessibility | Button already paints and animates disclosure state. The generic control UIA pattern dispatch returns no ExpandCollapse provider. | DxUi: expose confirmed disclosure state and ExpandCollapse commands coherently. Consumer owns persistence and task visibility. |
| Long paths / short body | Label supports wrapping and ScrollPanel has content height/offset and focus/input infrastructure. | Consumer: vertical role/name/location/metadata composition, exact copy value and body/actions regions; extend generic scrolling only for a reproduced gap. |
| Speed validation | Existing dialog is host-owned, with measured text helpers available. | Consumer: reflow on create, validation and DPI, keeping explanatory text separate from actions. |
| Issues | Grid, stable IDs and bounded rows exist. TaskDiagnosticEntry has no typed action/result correlation. | Consumer: typed diagnostic facts, summary priorities and selected detail. Generic Grid changes require a demonstrated defect. |

## First executable witness

Extend `Phase6_PopupSmokeResizeAndPause` to require Pause, Cancel and Speed controls with positive
rectangles and matching task/action hit targets once transfer graph/rate/ETA are ready. Correct
test-only layout observations so zero-height controls cannot count as visible. Production layout
is deliberately unchanged for the failing baseline. The fixture still exercises its existing
pause, resize, resume and cancellation sequence after this check passes.

Before the assertion, the fixture performs four warm-up paints and five rounds of 16 synchronous
owned-popup redraws, emitting `FileOps.Popup.ActionsWitness.RedrawBatchUs` with round and DPI.
This measures the entire synchronous redraw call, not display FPS; it does not replace the broader
resource/latency gates. Baseline and candidate use the identical fixture and profile.

Other pending witnesses: four direct actions across typed policies, scope checked-state/UIA,
French measurement overflow and 96-DPI validation overlap. The current change does not claim those
gates passed or P0 complete.

## Measurement status

The canonical library's common Release fixture was rebuilt and measured before library changes;
its raw five-round receipt is under `DxUi/Measurements/LocalizedAdaptiveLayout/2026-09-19/`.
That is independent offscreen evidence, not a consumer baseline or native presentation qualification.
Retain the consumer's focused runtime/performance artifacts before correcting the ordinary action row.
The final targeted localized-layout benchmark still needs its own unchanged-implementation run.

The first application build attempt was rejected because planning documents changed during its
source-attestation window. No successful receipt or runtime result was published. Freeze repository
edits during the rerun; do not bypass the source-identity check or describe that attempt as a test failure.

The rerun built successfully with no warnings/errors at `a212edec433d658748ae1273cba68633c86de8b2`
(the intervening commit contains only the approved execution-plan section). Its native witness failed
as expected: Pause, Cancel and Speed all had zero area and inactive hit targets. The five redraw batches
were 134593, 116594, 130145, 132722 and 133484 microseconds at 144 DPI; median 132722 microseconds.
See the [retained baseline](../../../TestRuns/4cb089111a23/FileOps/2026-09-19_101446-i26-actions-baseline/README.md).

The independent sandbox audit also reported 225 pre-existing run directories and three task log files
at the sandbox root. The three logs have been moved into its permitted `evidence/i26-ui` directory;
unrelated run directories are preserved. This audit is not a product-test pass or a clean final gate.
The current candidate reserves one ordinary action row independently of conflict rows; its runtime
result and paired performance comparison are pending. The other P0/L0 witnesses remain open.

The first candidate built with zero warnings/errors but its witness still failed. Cancel and Speed
now had positive rectangles. The test incorrectly required Pause while discovery was active (that
state offers Skip), and tested hit targets before waiting for the existing layout-motion flags to
settle. Retain this attempt under `Specs/TestRuns/4cb089111a23/FileOps/2026-09-19_102002-i26-actions-first-candidate`.
It is not accepted correctness evidence. The corrected fixture checks Skip during discovery and
Pause afterward, waits for layout/selected-card motion, and retains matching hit-target assertions.
Repeat its baseline with the production height repair temporarily absent, then reapply and qualify
the repair with the identical corrected fixture. Do not compare timings across fixture revisions.

The corrected baseline now reproduces the defect after layout settles: primary (Skip), Cancel and
Speed each report false/false, with discovery=true. Zero build warnings/errors; Setup/Cleanup pass.
The [v2 baseline](../../../TestRuns/4cb089111a23/FileOps/2026-09-19_102600-i26-actions-baseline-v2/README.md)
retains five redraw rounds (median 151136 microseconds), exact fixture SHA and build receipt.
The one-line height repair is restored for the identical candidate run; no product pass claimed yet.

## Accepted first slice

The corrected candidate passed all three declared outcomes (Setup, popup smoke, Cleanup), including
positive bounds/hit targets and pause/resize/resume/cancel. Baseline and candidate share the exact v2
fixture SHA and toolchain identity, x64 Debug and native 144 DPI. Median synchronous 16-redraw batch
duration was 151136 -> 133352 microseconds; this establishes no broader FPS/resource claim.
See [paired evidence](../../../TestRuns/4cb089111a23/FileOps/2026-09-19_103033-i26-actions-candidate-v2/README.md).

The [fresh harness pack](implementation/2026-09-19-ordinary-actions/README.md) contains 300 PNGs / 119
static French scenarios, with hashes and build provenance; the capture case passed. Both physical
monitors now report 144 DPI (independently queried), so fresh native 96-DPI qualification remains open.
Minimum-width footer text still clips. Other P0/L0 witnesses, adaptive changes and full gates remain.

The capture-only `-SkipBuild` attempt was correctly rejected because focused builds produce an
application-target receipt and that mode requires a full-solution receipt. Source/output hashes
still matched. The normal runner rebuilt and captured successfully. Do not bypass attestation.
