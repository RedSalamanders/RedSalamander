# Delete presentation pair; clipboard gate remains failed

Same-machine x64 Debug, qualified DxUi `78b3de3`, consumer `ecb391780` plus
recorded dirty snapshots. Final candidate source snapshot:
`1d0e0386e502c2c331cde041726979c7f705386e0c4250f63c910c216338c99e`.
Final full-solution build: zero warnings/errors. The preceding build exposed two
C4365 warnings in label-ID selection; the final build fixes them.

Identical gallery fixture SHA-256:
`6611c6abfa3123f9a8fc50858e2ca6a2bb23ad7196a2f09630f65ff14eeddebe`.
Baseline fails 0/1, zero skips: graph/disclosure present, accessible rate absent.
Candidate passes 1/1, zero skips at native 144/96 DPI and 480/760 DIP: no
graph/disclosure and an accessible rate descriptor. Reviewed 480-DIP frames keep
rate/count/path, progress and buttons separate. Ordinary expanded Delete shrinks
from 308 DIP including disclosure to 152 DIP; outer placement is preserved.
Open-discovery deletion rate and native AT readback are not qualified here.

Commands remains **FAILED: 2 passed, 1 failed, zero skips**. Progress contracts
and finished-task exclusion pass. Throughput disclosure fails later in its OLE
roundtrip at `Flush the owned rich-format snapshot before readback.` Restoration
readback also fails. This does not qualify clipboard or full interaction.
The next diagnostic must retain the exact publish/flush HRESULT.

Unmodified results/traces, aggregate/build provenance and capture manifests are
retained. Selected Delete redraw rows preserve original order; `selection.json`
binds raw digest, byte/line count and retained line numbers. Full inputs remain in
`C:/RedSalamander.Perf/evidence/i26-ui/deletion-graph-*-20260921/`.
Scratch-audit findings remain in aggregates; no clean-sandbox claim is made.

Five warm eight-redraw samples per case, microseconds:

| DPI / width DIP | Baseline median (min–max) | Candidate median (min–max) |
|---|---|---|
| 96 / 480 | 66602 (42994–90195) | 65753 (61229–72841) |
| 96 / 760 | 66469 (56447–76783) | 66593 (66220–66952) |
| 144 / 480 | 67304 (58802–71225) | 66578 (50117–82396) |
| 144 / 760 | 66431 (61428–74067) | 66752 (62761–70406) |

Medians remain within 1.3% with overlapping ranges. This small Debug pair does
not establish resource or tail-latency qualification. Graph absence is a
structural result, not a measured process-memory saving.
