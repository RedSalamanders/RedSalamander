# File Operations path keyboard: path-keyboard-initial

REJECTED: 2/3 pass, first repetition failed clipboard restoration. Exact-copy and focus-routing checks passed, but visual review also found a missing focus outline. The initial cleanup retained an OLE proxy without materializing its formats; the failing HRESULT was not logged, so delayed-format invalidation is a suspected cause, not a proven diagnosis.

Result: **2 passed, 1 failed, 0 skipped**. The build reports zero warnings/errors. Exact source/build identities, all selected redraw samples and capture populations are retained in the adjacent records.

The application harness uses its existing warning and desktop lease for the directed keyboard phase. It activates only the owned popup, then attempts to restore the previous foreground/cursor and preserves clipboard contents. The final fixture materializes advertised clipboard formats in an owned Shell data object before copying; an incomplete snapshot fails before mutation. It restores and flushes that object and checks both HRESULTs. This bounds the fixture to 256 formats and is not universal clipboard-provider qualification.

All application images come from WGC window capture. Raw data/logs remain under `C:/RedSalamander.Perf/evidence/i26-ui/path-keyboard-initial`. `metrics-compaction.json` binds each retained row to the original raw-file hash, line number and normalized LF digest. No selected failure/sample is discarded. Only trailing whitespace is normalized in runner.txt.

Focused Commands results do not replace Fresh Full, Release, native mixed-DPI or screen-reader qualification. General scrolling/selection of arbitrarily long detail is still open.
