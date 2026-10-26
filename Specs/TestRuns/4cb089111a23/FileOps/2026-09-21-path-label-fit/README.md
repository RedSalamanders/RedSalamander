# Native French path-label measurement pair

Same machine (`4cb089111a23`), x64 Debug, qualified DxUi `78b3de3`, consumer
`a0251c2361d8013d6cddc371b1b1b20e68d7367a` plus the recorded dirty candidate.
Both runs use the identical gallery fixture SHA-256
`1b81c40bfde3194cac54f90dd76537e9346ec96120e56e57e39aeed923b89950`.
Run IDs are `i26-path-labels-baseline-20260921` and
`i26-path-labels-candidate-20260921`. The runner aggregate records source/build
identities. Candidate source snapshot:
`b0655290f53f2964219259de4961c333ad56dffcba181c3c6bfc3ef1eb0cb863`.

Each variant retains the unmodified suite results/trace, aggregate, capture
scenario/display provenance and all `FileOps.Popup.PathLabels.*` metrics in
original order. `selection.json` binds the full raw stream's SHA-256, byte/line
counts, retained original line numbers and retained stream digest. Full inputs
remain under `C:/RedSalamander.Perf/evidence/i26-ui/path-labels-<variant>-20260921/`.

The baseline failed 0/1 (zero skips); the candidate passed 1/1 (zero skips).
Painted labels are checked against their actual DirectWrite format and paint
rectangles, independently of the production width cache. At 144 DPI deletion
slack changed from -12.970 to +1.030 pixels. At 96 DPI deletion changed from
-11.305 to +1.695 and verification from -1.030 to +1.970 pixels. Verification
at 144 DPI remains +2.668. Both 480 and 760 DIP widths passed in the candidate.

Warm verification repaint timings are five samples of eight synchronous redraws
per configuration, after three warm redraws. Values below are microseconds.

| DPI / width DIP | Baseline median (min–max) | Candidate median (min–max) |
|---|---|---|
| 96 / 480 | 74238 (58025–99286) | 69325 (55895–75911) |
| 96 / 760 | 71245 (60358–86083) | 63055 (51295–81976) |
| 144 / 480 | 66621 (45482–87571) | 71808 (48691–101092) |
| 144 / 760 | 74318 (48281–81554) | 60353 (43876–105478) |

There is no consistent slowdown across this pair; the 144/480 median increased
7.8%, while the other medians decreased and all sample ranges overlap. Twenty
noisy Debug batches do not establish tail-latency, whole-product performance or
resource non-regression. Baseline/candidate scratch audits respectively reported
229/230 pre-existing or retained-run findings; no clean-sandbox claim is made.
These focused builds target the application project, not the full solution.
