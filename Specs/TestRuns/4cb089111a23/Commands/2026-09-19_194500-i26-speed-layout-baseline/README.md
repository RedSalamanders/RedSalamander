# Speed dialog — original baseline

`speed-layout-baseline` fails 0/1, zero skips, with zero build warnings/errors and a clean sandbox
audit. The French constrained-client witness finds help bottom 132 DIP and buttons top 92 DIP in
a 300x142-DIP client. All four states and 20 eight-redraw samples are retained despite that expected
failure. Median is 66,527 us; nearest-rank p95 is 102,916 us.

The [four-run comparison](../2026-09-19_194500-i26-speed-layout-candidate-repeat/README.md) binds exact
workspace/receipt/fixture identities and preserves the initial unfavorable timing flag and repeats.
Complete metrics/results are retained here. Only trailing runner whitespace is normalized. Raw
artifacts remain under `C:/RedSalamander.Perf/evidence/i26-ui/speed-layout-baseline*` and
`S:/RedSalamander.Perf/evidence/speed-layout-baseline`.
