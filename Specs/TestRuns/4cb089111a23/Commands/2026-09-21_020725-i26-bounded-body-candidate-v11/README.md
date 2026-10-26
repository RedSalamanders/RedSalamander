# Bounded conflict comparison: candidate-v11

Result: **6 passed / 0 failed / 0 skipped**, x64 Debug.
Source/build snapshots match. Both runs exercise comparison and typed-policy guards three times;
the candidate adds bounded keyboard/UIA assertions at the end of comparison coverage. Inspection,
close/Escape, completed navigation and pane navigation run separately as a 15-case regression set.
The first auto-expanded fixture attempt is retained separately and is not the short-window baseline.

The fixture delivers the resize message sequence, asserts a 520-DIP outer height, and supplies long
French source/destination folders. It requires positive, fully visible decisions and complete scope;
wheel input must move comparison content while the filename and decisions remain stationary, and
the admitted engine conflict must remain unresolved. Actual captures are produced by the harness.

Common/differing-name redraw batches use the same 480/640/760-DIP scene inputs (45 samples
each), but the baseline predates the final placement/device/geometry assertions. V11 common-name
p95 is 1.818 s, with multiple hosted-paint stalls. This is not a performance pass. Existing per-frame instrumentation additionally retains the 520-DIP window's layout/draw/total
timings using ClientPixels markers; resize/scroll transients, unequal frame counts and the candidate's
additional keyboard scene are explicit. These short-window samples are diagnostic and unpaired;
they are not steady-state CPU/GPU/resource or Release qualification.

Raw: `S:/RedSalamander.Perf/evidence/i26-bounded-body-candidate-v11-20260921`. Log: `C:/RedSalamander.Perf/evidence/i26-ui/bounded-body-candidate-v11.log` (SHA-256 `2c5704d68e48b032d1a7a6abdc13d159a0767ec6319a9ff68a12a1d32384c7f3`).
Every selected metric retains its original source line and source-file hash. Baseline failures remain
failures; candidate acceptance requires review of the actual before/after images and regression results.
