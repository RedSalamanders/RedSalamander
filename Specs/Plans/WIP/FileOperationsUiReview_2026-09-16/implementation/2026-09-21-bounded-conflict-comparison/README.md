# Bounded conflict comparison — 2026-09-21

[Actual harness gallery](gallery.html) · [Capture provenance](capture-manifest.json) · [Measurements](paired-measurements.json)

The earlier display pushed decisions below a short window. The candidate keeps the common filename,
conflict state, inspection feedback, complete scope caption and decision rows visible. The stacked
folder comparison scrolls independently; wheel, keyboard, paint and UIA use the same viewport.
A replacement prompt rebuilds its retained panel even when its task ID stays unchanged.

A qualified common filename appears once. Differing names each have one bold row, with parent-only
locations below. Unknown namespaces preserve the exact-path fallback. Source stays above destination.
Inspection navigates to the admitted item without answering the conflict; its result stays above scope.

## Verification

- Corrected baseline: **3 passed / 3 failed**; each comparison loses its decisions below the viewport.
- Final candidate-v11: **6/6**, plus inspection/close/reopen/navigation **15/15**; no skips or disk-audit issues.
- Full x64 Debug build: zero warnings/errors. Source/localization Pester: **199/199**.
- Real 520-DIP outer height, positive complete decision bounds, stationary filename/actions, wheel movement,
  inspection focus reveal, Tab to scope and unresolved engine state are asserted.
- Pixel probes require the fixed explanation to retain at least 95% of its text pixels after keyboard
  reveal and require a visible scope-focus stroke. Input uses the existing warning/desktop lease and
  restores cursor/focus. Pixel checks supplement geometry; the earlier visual concern alone did not establish clipping.
- The final captures were directly reviewed at 480/640/760 DIP, native 144 DPI, including before/after
  scrolling, keyboard-before/after, differing names, provider fallback and inspection feedback.

Intermediate failures remain in raw evidence. The first auto-expanded fixture was rejected as a short
window witness. Product fixes address a stale retained panel and duplicate inspection paint; the overlap
audit clips only inspection content to its actual viewport, preserving full decision/UIA assertions.
V7 and V10 found a monitor-handle mismatch. Final verification checks display device, monitor/work
rectangles and DPI instead of opaque handle lifetime; changing any required display fact still fails.
V10 also found an abandoned initial build-scratch directory, now preserved under evidence; it was not
an independently running application. Regression V10 itself passed 15/15.

## Measurements and limits

45 ordinary redraw batches per name profile; same width/scenario inputs, microseconds median/p95:

| Scene | Baseline | V6 unchanged confirmation | Final candidate-v11 |
|---|---:|---:|---:|
| Common filename | 66550 / 68069 | 66512 / 68481 | 66557 / 1817984 |
| Differing filenames | 66606 / 66906 | 66532 / 68594 | 66542 / 67087 |

The first V6 candidate crossed 144→96→144 DPI and had 2.06 s/1.17 s batches. It is **invalid as a
matched timing sample**; DPI transition is a confound, not a proved sole cause. All samples remain.
Its extra `96dpi` filename sampled DPI before resize and is labelled diagnostic, not density acceptance.
V11 still has a common-name p95 of **1.818 s**, despite unchanged display facts. Existing instrumentation
places the stalls in `HostedPaintUs` (roughly 250–520 ms per affected frame), not comparison layout.
That sample is **not a performance pass**. Next: route the pinned library's existing frame-stage sink
into the harness to distinguish render/update from swap-chain presentation. Do not rebaseline or
attribute the stall to DPI alone.

Final placement uses a work-area inset and asserts display facts per frame; the baseline predates these
fixture checks. These bounded redraw timings do not establish final non-regression. Short-window frame
samples are unpaired: extra keyboard coverage and unequal visible content/frame counts remain explicit.

Final steady-state CPU/GPU/resources, Release, native mixed DPI/192 DPI, AT/hover, multiple oversized
cards and removed-action/reflow acceptance remain I26 gates. The 760-DIP fixture deliberately leaves
space below the card. The inspection-result fixture still has an English caption despite French content.

## Receipts

Final: `Specs/TestRuns/4cb089111a23/Commands/2026-09-21_020725-i26-bounded-body-candidate-v11`.
Earlier baseline/V6/V6-confirm receipts remain under `Specs/TestRuns/4cb089111a23/Commands/`.
Raw runs: `S:/RedSalamander.Perf/evidence/i26-bounded-body-*-20260921`.
Logs/scripts: `C:/RedSalamander.Perf/evidence/i26-ui`.

Mandatory Full qualification is still failed. This component does not close I26.
