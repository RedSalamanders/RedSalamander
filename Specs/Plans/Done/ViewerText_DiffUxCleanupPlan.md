# ViewerText Diff UX Cleanup Plan

Status: Implemented

## Goals

- align parsed diff presentation more closely with the intended editor-style diff surface,
- simplify diff mode routing so raw text is just the normal text view,
- make hidden unchanged ranges actionable from the surface,
- validate that the cleanup does not regress the protected viewport rehydrate/backtrack and combo-sync paths.

## Checklist

- [x] `View / Text` switches diff documents to raw text and the header mode button cycles `Diff -> Raw -> Hex`.
- [x] Parsed diff rows use the normal text background for unchanged content instead of dedicated context row fills.
- [x] Leading `+` / `-` diff markers render dimmed against the current theme background.
- [x] Absent panes / gaps render with subtle hatched background treatment instead of flat empty fills.
- [x] Parsed diff hunk anchor rows are no longer rendered as visible `@@ ... @@` document rows.
- [x] Hidden-context banners are restyled and can be clicked to reveal unchanged text.
- [x] Parsed diff horizontal scrolling is validated for both `side-by-side` and `inline` presentations after wrapping is disabled, including presentation switches.
- [x] Shared viewer combo hosts expand to show popup rows on click, collapse on `Escape`, and close without detaching/crashing across `ViewerPE`, `ViewerWeb`, `ViewerImgRaw`, and `ViewerText`.
- [x] Specs are updated for the new menu/mode contract and row rendering rules.
- [x] Runtime tests cover menu/mode routing, hidden-banner click-to-reveal, non-anchor parsed presentation, dual-mode horizontal scrolling, and shared viewer combo-host popup activation.
- [x] Perf selftest evidence is refreshed, with explicit focus on viewport rehydrate/backtrack and combo-sync behavior.
- [x] Post-plan tuning keeps hidden-banner perf dispatch reliable after hunk navigation and trims preserve-viewport combo/header churn during diff rehydrate/backtrack rebuilds.

## Notes

- This is a follow-up cleanup pass after [ViewerText_DiffViewerPlan.md](ViewerText_DiffViewerPlan.md) and [ViewerText_DiffPolishAndPerfPlan.md](ViewerText_DiffPolishAndPerfPlan.md).
- Raw text remains supported as a session presentation and as `diffAutoOpenMode`, but it should no longer need its own dedicated diff submenu entry.
- The latest same-machine post-plan perf candidate is [viewer_text_diff_perf_metrics.json](../../TestRuns/4cb089111a23/Commands/2026-04-05_112653/perf/viewer_text_diff_perf_metrics.json), compared against [viewer_text_diff_perf_metrics.json](../../TestRuns/4cb089111a23/Commands/2026-04-04_152319/perf/viewer_text_diff_perf_metrics.json) after the combo-host close-path and dual-mode horizontal-scroll fixes.
- That same-machine comparison improved the large parsed side-by-side open (`349105 us -> 290126 us`), theme-switch repaint (`11024 us -> 10538 us`), side-by-side scroll repaint (`9453 us -> 8692 us`), resolved expanded open (`104630 us -> 92935 us`), hunk jump (`8086 us -> 7761 us`), and expand context (`24106 us -> 20773 us`) while keeping the bounded referenced-byte contract unchanged at `32768`, `65536`, and `0` reread on backtrack.
- The main caveat on the latest candidate is that the unresolved-placeholder scenario regressed slightly (`92454 us -> 98088 us`) and resolved viewport rehydrate/backtrack are mixed (`165913 us -> 167143 us`, `80331 us -> 81092 us`), so the perf result is a net win rather than a clean across-the-board improvement.
