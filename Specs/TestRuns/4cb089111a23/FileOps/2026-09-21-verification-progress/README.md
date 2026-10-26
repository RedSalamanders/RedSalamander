# Verification whole-task progress correction

Same machine, x64 Debug, qualified DxUi `78b3de3`. Candidate is consumer
`27a1670ce` plus source snapshot
`c6dae2f53545ca930d03942b85347ab26ebbd11942036579f062a1155b075c15`.
The full-solution build passed and the governed runner reused its verified receipt.
Focused Commands passed 2/2, zero skips: progress contracts and finished-task
exclusion. The new assertions cover proof/transfer denominators, mixed weights,
open totals, conflict exclusion, pause, completed proof and saturated byte counters.
The native French gallery passed 1/1, zero skips. Both 480-DIP verification
captures were reviewed at native 144/96 DPI; footer and task now both show 75%.

Baseline is the [prior path-label candidate](../2026-09-21-path-label-fit/candidate/).
The entire gallery fixture is unchanged, SHA-256
`1b81c40bfde3194cac54f90dd76537e9346ec96120e56e57e39aeed923b89950`.
All claim-bearing path-label/redraw metrics retain original order and are bound
to the raw stream's digest, size and line indices by `selection.json`.
Full raw runs remain under
`C:/RedSalamander.Perf/evidence/i26-ui/verification-progress-*-20260921/`.
Unmodified per-case results, traces, aggregate/build provenance and scenario/display
manifests are retained here. Scratch audit findings remain recorded in the aggregates.

Five warm eight-redraw samples per case, microseconds:

| DPI / width DIP | Baseline median (min–max) | Candidate median (min–max) |
|---|---|---|
| 96 / 480 | 69325 (55895–75911) | 72537 (45410–80557) |
| 96 / 760 | 63055 (51295–81976) | 66581 (44832–85677) |
| 144 / 480 | 71808 (48691–101092) | 55233 (51427–80974) |
| 144 / 760 | 60353 (43876–105478) | 66643 (59391–67962) |

Three medians rose (4.6%, 5.6%, 10.4%) and one fell (23.1%); all ranges overlap.
This small Debug sample is directional and does not establish performance,
tail-latency or resource non-regression. Final P5 qualification remains open.
