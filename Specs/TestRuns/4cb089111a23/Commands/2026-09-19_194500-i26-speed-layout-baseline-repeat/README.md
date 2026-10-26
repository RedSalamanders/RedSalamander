# Speed dialog — repeated baseline

`speed-layout-baseline-repeat` reproduces the expected French overlap in all three repetitions:
0 passed, 3 failed, zero skips, zero build warnings/errors and a clean sandbox audit. Its workspace
is identical to the original baseline. All 60 ordered eight-redraw samples remain intact: median
66,509.5 us, nearest-rank p95 121,362 us. Baseline p95 variability is itself evidence; these repeats
supplement the original baseline and do not replace it.

The [four-run comparison](../2026-09-19_194500-i26-speed-layout-candidate-repeat/README.md) retains
both pairs and exact source/receipt/fixture bindings. Raw artifacts remain under
`C:/RedSalamander.Perf/evidence/i26-ui/speed-layout-baseline-repeat*` and
`S:/RedSalamander.Perf/evidence/speed-layout-baseline-repeat`. Runner whitespace alone is normalized.
