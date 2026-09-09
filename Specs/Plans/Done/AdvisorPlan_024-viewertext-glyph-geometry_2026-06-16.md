# Advisor Plan 024 - ViewerText DirectWrite geometry

> **NON-NORMATIVE COMPLETE RECORD.** Archived from root `plans/024-viewertext-glyph-geometry.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** COMPLETE
- **Archived:** 2026-08-25
- **Original path:** `plans/024-viewertext-glyph-geometry.md`
- **Live owner:** Specs/Plans/Done/Operation_Farsight_ViewerPluginsRemediation_ExportSafetyDecodeReliabilityWebSecurityAndTextGeometry_2026-06-16.md
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 024: ViewerText — DirectWrite-accurate geometry with bounded sparse wrapping

## Status

- **State**: DONE — implementation, focused geometry/sparse proof, consolidated archive, and repository-wide Full gate green (2026-07-12)
- **Priority**: P2 (wrong selection/caret/click behavior and huge-line responsiveness)
- **Effort**: L
- **Risk**: MED–HIGH (geometry, navigation, wrapping, cache, and async streamed state)
- **Category**: correctness / responsiveness / bounded memory
- **Planned at**: commit `b274022d9`, 2026-06-16

## Historical defect

Text mode mapped UTF-16 buffer columns to pixels with `column * charW`. DirectWrite does not: tabs,
proportional/CJK glyphs, combining sequences, and surrogate pairs shape into clusters with real
advances. Selection, search rectangles, caret placement, and hit testing therefore drifted and the
caret could enter a surrogate pair. The original plan also assumed that per-visible-line layout alone
was sufficient; the closeout audit found that wrapping a multi-million-code-unit line still
materialized unbounded visual-row metadata and could do expensive streamed/index work on the UI thread.

## 2026-07-12 implementation reconciliation

### One geometry source of truth

Visible text segments use `IDWriteTextLayout`. Drawing and all content geometry derive from the same
layout/cluster model:

- selection/search rectangles use `HitTestTextRange`;
- caret positions use `HitTestTextPosition`;
- mouse mapping uses `HitTestPoint`;
- tabs use the configured incremental tab stop;
- horizontal extent comes from measured layout width;
- left/right movement and selection edges stay on valid DirectWrite/UTF-16 cluster boundaries.

Vertical/page navigation preserves the preferred X coordinate and keeps a bounded exact inverse
history, so Down→Up and PageDown→PageUp restore the original caret when content is unchanged. Mouse
drag uses the message's `MK_LBUTTON` state rather than sampling global key state.

### Bounded layout and sparse wrap state

The layout cache defaults to 64 entries / 2 MiB, has deterministic LRU-style eviction telemetry, and
is invalidated on width/wrap/text-window/encoding/DPI/theme/font/device changes. Huge wrapped content
stores sparse per-logical-line summaries and derives only the requested viewport rows (hard ceiling
4,096). A bounded checkpoint cache stores, for each split row, independent plain/left/right cursors;
it never substitutes `max(left,right)` for unequal pane progress. Sparse anchors are explicit and need
not be numerically contiguous.

In sparse mode `textVisualLineCountExact == false`: the reported coordinate span is suitable for
scroll mapping but is not misrepresented as an exact materialized row count. Every code unit remains
reachable through logical navigation, scrolling, selection, and search.

### Streamed/terminal work

Expensive streamed read/decode/index work is module-pinned, identity-bound, generation-checked, and
terminal on every current submit/worker/post failure. It never joins the UI thread. Exact reader and
decode boundary behavior are shared with the ViewerText reliability slice.

## Reconciled scope

The implementation spans `ViewerText.Text.cpp`, `ViewerText.cpp`, `ViewerText.h`, shared Debug
contracts, ViewerPETests, source contracts, instrumentation, resources, and the authoritative
ViewerText spec. The original “only `ViewerText.Text.cpp` (+ header)” criterion is obsolete as of
2026-07-12 because bounded sparse wrapping and deterministic proof are part of the correction.

Hex rendering remains outside the glyph-geometry implementation, except for shared terminal/exact-
reader contracts handled by the companion ViewerText slices.

## Focused proof

Focused ViewerText regressions cover:

- tabs, proportional Latin, CJK, combining sequences, emoji/non-BMP clusters, search rectangles,
  selection, clicks, and left/right navigation;
- wrapped-segment width (including an unavoidable single over-width cluster);
- exact Down→Up and PageDown→PageUp caret restoration;
- synthetic and real mouse drag selection;
- bounded entry/byte cache behavior and eviction;
- huge-line sparse activation without per-row materialization;
- unequal split-pane advancement with independent left/right cursors and exact reverse navigation;
- rapid streamed navigation, stale rejection, terminal failures, close, and unload.

`TestViewerTextAsyncOpenAndUtf8HexTerminalContracts` and
`TestViewerTextDiffModesAndPlaceholders` are focused-green; scoped ViewerText/ViewerPETests builds are
zero-warning.

Consolidated proof is archived at
`Specs/TestRuns/4cb089111a23/Viewers/2026-07-12_105903_farsight_viewer_closeout/`: all 13 focused
cases passed with zero conditional skips and 1,426 performance metric records, including the async
terminal and diff/sparse ViewerText cases.

## Done criteria

- [x] Selection, caret, search, hit-test, and horizontal extent use DirectWrite layout metrics.
- [x] Tabs, CJK/proportional text, combining clusters, and non-BMP characters align in focused tests.
- [x] Caret/selection never split a surrogate pair or shaped cluster.
- [x] Vertical/page navigation preserves preferred X and exact inverse history within its bound.
- [x] Layout cache entry/byte caps and deterministic invalidation/eviction are enforced.
- [x] Huge wrapped lines use sparse summaries and viewport-only derived rows.
- [x] Split panes checkpoint independent left/right cursors; unequal panes navigate correctly.
- [x] Sparse telemetry distinguishes coordinate span from exact row count.
- [x] Streamed work is non-blocking, latest-wins/identity-bound, and terminal on current failures.
- [x] Scoped zero-warning builds and focused geometry/stream/sparse regressions pass.
- [x] Expanded shared/test/resource/spec scope is recorded; the old narrow-file criterion is retired.
- [x] `plans/README.md` reflects this reconciliation.
- [x] Consolidated ViewerText/Farsight metrics archive is recorded/refreshed.
- [x] `.\Tools\Run-AllTests.ps1 -Suite Full` passes in the final closeout workspace (`Specs/TestRuns/4cb089111a23/Continuation/2026-07-12_205636_farsight_full_green/`).

## STOP conditions

- Any content geometry falls back to fixed code-unit columns.
- Sparse navigation assumes viewport anchors are contiguous or merges split-pane cursors.
- Cache/sparse metadata scales with total wrapped rows rather than the viewport/bounds.
- A streamed completion can mutate a stale/recycled window or leave a current request loading.

## Maintenance notes

The invariant is broader than “use DWrite to draw”: the same shaped layout/cluster data must govern
draw, selection, search, caret, hit testing, navigation, and scroll extent, while caches and sparse
metadata stay explicitly bounded.
