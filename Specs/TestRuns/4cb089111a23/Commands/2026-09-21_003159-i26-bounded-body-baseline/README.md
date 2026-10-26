# Bounded conflict comparison: baseline

Result: **3 passed / 3 failed / 0 skipped**, x64 Debug.
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

Raw: `S:/RedSalamander.Perf/evidence/i26-bounded-body-fixed-fixture-baseline-20260921`. Log: `C:/RedSalamander.Perf/evidence/i26-ui/bounded-body-fixed-fixture-baseline.log` (SHA-256 `0bf9449969b53efb797a650d4727abce09252c0df1471599b0385981a87f635b`).
Every selected metric retains its original source line and source-file hash. Baseline failures remain
failures; candidate acceptance requires review of the actual before/after images and regression results.
