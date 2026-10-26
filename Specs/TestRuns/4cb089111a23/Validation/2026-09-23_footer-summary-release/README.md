# Wrapped French footer summary: Release evidence

The footer reserves the measured summary height before positioning commands. One
current DirectWrite layout is shared by measurement and painting; unchanged redraws
reuse it. Minimum sizing also measures the summary at the supported 480-DIP width
using the same layout, with no text allocation in WM_GETMINMAXINFO.

## Intermediate attempt

Nine explicit candidate sources are recorded under `raw/intermediate-gallery`.
Build receipt `5ce1c4aac46b271cbb2c5c2f495f177c5f9390571b3e9611a035f485a540ac98`
completed with two LONG-to-float diagnostic warnings, later corrected.
Run `20260923T103616Z-67904-46457ccb3f714e938e01bf715db184f7` passes all eight
new footer geometry checks, but the overall case fails the confirmation pixel
witness: 0 passed, 1 failed, zero audit findings. This is not a gallery pass.

| Physical DPI | Client width | Lines | Required / allocated height | Modes |
| --- | --- | --- | --- | --- |
| 144 | 480 DIP | 2 | 43.893 / 44 px | Footer-only and full |
| 144 | 760 DIP | 1 | 21.946 / 30 px | Footer-only and full |
| 96 | 480 DIP | 2 | 29.262 / 30 px | Footer-only and full |
| 96 | 760 DIP | 1 | 14.631 / 20 px | Footer-only and full |

The checks require the actual measured text to fit, every visible command to remain
below the summary and inside the client, and Queue to have an active hit target.
The 480-DIP fixture must actually wrap; a one-line fallback cannot pass.
Full unmodified logs/images/metrics are preserved at
`C:/RedSalamander.Perf/evidence/i26-ui/footer-summary-wrap-20260923/`.

## Final candidate and limits

The final candidate adds the minimum-sizing measurement and fixes both warnings.
Its Release build passes with zero warnings/errors in 291,435 ms, receipt
`ce4b906c61a6d1cf19f1b5d66fad72ea5789556a22f5cc0f99b43850bbd7c400`, source snapshot
`6f6a6b0faf385d16ed815b38aef5c392338d4eabb812d8f3475de4276b9517cc`.
The nine source-input hashes match the primary and isolated working trees. Spec inventory
has zero blocking findings. Fresh PR run
`20260923T104358Z-68828-0c7698080a4b473294d3edc81f10d070` passes **13 executed/promoted
entries, 824 reported passes, zero failures/skips/audit findings**, exit 0, in 421,951 ms.
There are no reused or provisional outcomes and no quarantine retries. The count contains
812 Pester cases and twelve native/script entry outcomes, not every detailed native assertion.
The updated ProductUi footer regression passes. Full raw receipts are retained under `raw/final/`;
external runner logs and entry evidence are at
`C:/RedSalamander.Perf/evidence/i26-ui/footer-summary-final-20260923/`.
This is final-candidate PR coverage. Complete confirmation captures, Full, native focus/UIA/AT,
192-DPI, combined DxUi pin and resource/performance acceptance remain open.
The unchanged application pin is `78b3de389a189c7f86f611787e0489fb6d474218`.

[Active plan](../../../../Plans/WIP/UI_FileOperationsReview_2026-09-16.md).

## Final-source gallery and measurements

Gallery `20260923T105229Z-28376-e22cb67e77bc4be8be21e4219f9de2ca` confirms the same
eight summary fit checks on the final source. It still fails the confirmation witness
(0/1 overall, audit zero). It retains 285 PNGs: 276 shared originals byte-identical to
baseline, one rejected black shared confirmation, eight new summary scenes. There are
27 missing baseline images after failure, explicitly listed in the comparison.
[Actual captures and review](../../../../Plans/WIP/FileOperationsUiReview_2026-09-16/implementation/2026-09-23-footer-summary-release/README.md).

The shared delete/verifying fixtures preserve five eight-redraw samples per width/DPI.
Medians remain presentation-paced near 64.25–66.69 ms. The largest median increase is
64.916 to 66.627 ms for verifying at 96 DPI/480 DIP. Whole-gallery counts differ because
of new scenes and early failure, so their aggregates are not a controlled comparison.
The new summary cache records 179 misses, median 21 microseconds, p95 99 microseconds.
All focused samples and full external metric identities are retained. This one serial
pair does not establish a memory, CPU, steady-state throughput or resource gate pass.

Canonical FileOps is running with the corrected alternate-volume option. Its outcome
will be recorded separately; there is no overall FileOps/Full claim yet.
