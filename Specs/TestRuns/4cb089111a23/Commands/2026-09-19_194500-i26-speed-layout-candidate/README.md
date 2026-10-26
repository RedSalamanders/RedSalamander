# Speed dialog — first candidate

`speed-layout-candidate` passes 1/1, zero skips, with zero build warnings/errors and a clean sandbox
audit. The French initial/constrained/validation/corrected geometry and existing UIA/value and
nested-shutdown assertions pass. This candidate predates the final minimized/empty-client guard.

Its 20 eight-redraw samples have median 66,657.5 us and p95 118,309 us. The initial p95 increase of
about 15% is retained as a flag, not discarded. See the [follow-up comparison](../2026-09-19_194500-i26-speed-layout-candidate-repeat/README.md).
Full metrics/results and exact build/workspace identities are retained; runner whitespace alone is
normalized. Raw artifacts remain under `C:/RedSalamander.Perf/evidence/i26-ui/speed-layout-candidate*`
and `S:/RedSalamander.Perf/evidence/speed-layout-candidate`.
