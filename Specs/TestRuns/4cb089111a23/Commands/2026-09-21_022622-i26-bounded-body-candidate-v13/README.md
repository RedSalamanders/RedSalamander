# Retained focus failure V13

**5 passed / 1 failed / 0 skipped**, three repetitions, clean disk audit, zero build warnings/errors.
Repetition one lacks a painted scope focus ring. Explanation pixels 304/304; focus pixels 0.
Before/after WGC: keyboard modality true, same scope automation ID, owned foreground true,
cursor outside the window at (3838,2), scope top unchanged at 452 px. The failure image is retained
here, separately from later successful repetitions. No user interruption or display movement is
inferred from these observations. The host control's actual focus flag, reduced-motion policy and
paint state still need diagnosis. Existing native screenshot probes remain strict.

The fixture now stores all PNGs under conflict-comparison/repeat-N and records pre/post capture
state. The frame-stage sink exposes update/render/presentation through the existing perf stream.
Selected metrics retain original source lines and hashes; redraw/frame samples do not establish
paired final resource or Release acceptance. The V12 warning is corrected in this source.

Raw: S:/RedSalamander.Perf/evidence/i26-bounded-body-candidate-v13-20260921.
Do not use a later passing repetition to erase this failure or close the focus gate.
