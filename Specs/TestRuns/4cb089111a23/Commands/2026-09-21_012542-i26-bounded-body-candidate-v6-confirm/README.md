# Bounded conflict comparison: candidate-v6-confirm

Result: **6 passed / 0 failed / 0 skipped**, x64 Debug.
Source/build snapshots match. Both runs exercise comparison and typed-policy guards three times;
the candidate adds bounded keyboard/UIA assertions at the end of comparison coverage. Inspection,
close/Escape, completed navigation and pane navigation run separately as a 15-case regression set.
The first auto-expanded fixture attempt is retained separately and is not the short-window baseline.

The fixture delivers the resize message sequence, asserts a 520-DIP outer height, and supplies long
French source/destination folders. It requires positive, fully visible decisions and complete scope;
wheel input must move comparison content while the filename and decisions remain stationary, and
the admitted engine conflict must remain unresolved. Actual captures are produced by the harness.

Common/differing-name redraw batches remain paired on identical 480/640/760-DIP scenes (45 samples
each). Existing per-frame instrumentation additionally retains the 520-DIP window's layout/draw/total
timings using ClientPixels markers; resize/scroll transients, unequal frame counts and the candidate's
additional keyboard scene are explicit. These short-window samples are diagnostic and unpaired;
they are not steady-state CPU/GPU/resource or Release qualification.

Raw: `S:/RedSalamander.Perf/evidence/i26-bounded-body-candidate-v6-confirm-20260921`. Log: `C:/RedSalamander.Perf/evidence/i26-ui/bounded-body-candidate-v6-confirm.log` (SHA-256 `35ed2ef2d9e1caa71c40a4a986c763d1836c7a21be471e8c09c2f9b3141967e1`).
Every selected metric retains its original source line and source-file hash. Baseline failures remain
failures; candidate acceptance requires review of the actual before/after images and regression results.
