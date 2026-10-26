# Throughput UI — first candidate and paired diagnostic

Run `20260919T141917Z-25120-9f5f84b1ff1a4ea6910def99acb23a8e`, x64 Debug, passes
`cmd_pane_fileops_popup_throughput_disclosure` 1/1 with zero skips and zero disk-audit issues.
Receipt `a10e555720fe23e715a49607cc662b79fb837348aa42ef9d81e3cb041c786708` binds the build;
zero warnings/errors. This run precedes the additional documentation captures.

The same-machine measured section matches the [baseline](../2026-09-19_160946-i26-throughput-ui-baseline/README.md):
one running-copy fixture, reduced motion, five eight-redraw batches after three warm-up paints.
Expanded values: 373680, 374754, 373490, 378155, 346268 us; median 373680 us versus
374500 us baseline (-0.22%). Collapsed: 373160, 376212, 374698, 361064, 375552 us;
median 374698 us. No material improvement is established. These synchronous Debug timings are
not presented-frame pacing or a Release resource regression gate. Later Debug screenshots reveal
missing legacy rendering, so the timings cannot establish visual correctness or complete-render cost.

The case verifies graph removal, summary/actions retained, unchanged outer rectangle, queued save,
real file reload, independent live cards, remembered defaults and compact/footer-only suppression.
Actual external hot reload, failed save, keyboard/UIA and both placement keys remain open.

Archived runner output has trailing whitespace normalized; unchanged full logs remain in the governed local evidence folder.
