# Frame-stage diagnostic run V12

**5 passed / 1 failed / 0 skipped**, three repetitions. The second keyboard capture retained
304/304 explanation pixels but reported zero focus-outline pixels. Its PNG was subsequently
overwritten by repetition three; that last image must not be presented as the failure witness.
V13 assigns each repetition a separate capture directory. Source/build snapshots match; build
had one C5246 warning in the new test initializer, corrected in V13. No disk-audit issues.

The fixture uses the pinned library frame-stage metric sink, verifying all four stages were
observed. 880 records per stage: presentation max 5389 us, render max 286784 us. The repeated
250–520 ms hosted-paint stalls from V11 did not recur in this run. An isolated render stall and
all redraw batches remain in evidence; neither a performance pass nor an OS-occlusion root cause
is inferred. Frame metrics and selected ordinary/short-window measurements retain original
source line numbers and source hashes in metrics-compaction.json. Short-window samples remain
unpaired; later fixture changes prevent treating these as identical baseline bytes.

Raw: S:/RedSalamander.Perf/evidence/i26-bounded-body-candidate-v12-20260921.
Stage analysis: stage-analysis.txt. Original build/runner logs remain in
C:/RedSalamander.Perf/evidence/i26-ui. Display-only archived logs may trim whitespace;
native UTF-16 traces and metric/JSON data are preserved.
