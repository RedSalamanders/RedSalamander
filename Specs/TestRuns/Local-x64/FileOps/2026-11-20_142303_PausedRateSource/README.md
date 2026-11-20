# Paused-rate Source study: qualitative A/B evidence

This derived package supports state correctness and bounded work in one controlled
local x64 Release case. A remains **FAILED** (native 1, runner 1; two passed and
one failed). B's filtered native case **PASSED** (native 0; three passed, none
failed or skipped), while runner exit 2 leaves repository coverage
**NOT_EVALUATED**. It establishes no Full, ARM64, desktop, merged-main or complete
repository acceptance.

The unchanged test and instrumentation ran against the reviewed five-hunk
inverse in three product files (A source c88af9df) and the corrected B2 source
(B source 736eff83), both based on ebebc14b. Each build has its own exact canonical
receipt and generated SDK archive/module identities. The Source recipe controls
pinned implementation and inputs, not cross-build generated archive equality.
The prospective common closure contains 35,584 files and 199 command tlogs,
including 465 retained producer read-witness paths. Those witnesses are not a
claim to enumerate every compiler input. The outer environment observation and
derived public binding remain separate; the inner compiler environment was not
observed completely.

A's fresh nonzero rate remained 3031.076581576027 after pause with no telemetry
timer and zero rate collections. B's fresh nonzero rate changed from
3049.125418682546 to zero, also before another telemetry timer, with one
collection. The separate repeated-pause ordering control deliberately held a
ZERO second rate: A had zero collections, one window sample and reset=false;
B had one collection, one window sample and reset=true. This is not a second
nonzero-rate witness.

Original settling (600 ms), requested measurement (2200 ms), and caps (3 timer
ticks, 4 popup task-snapshot calls, 4 owner ticks) remain unchanged. A observed
2208 ms and counters 2/2/2; B observed 2209 ms and counters 2/2/2. Both were within
the original bounds. B re-armed its 100 ms cadence in 2 ms. The trace label
"renders" counts popup task-snapshot calls, not monitor-presented frames.
**Neither half reproduced the earlier ten-render symptom.**

There is no statistical timing, speedup, speed ratio, executable-code-equivalence
or general performance claim. A CPU observation of 0 and B's 15,625 microseconds
are individual quantized observations, not a benchmark comparison. Diagnostic
observer rows add overhead and are emitted after measurement.

The selected actual metrics retain their payload values, original row order,
line ordinals and decoded-line digests. Native commit fields were EMPTY and have
not been filled. Separate source bindings carry the source and receipt identity.
CompletedCopies.durationUs is the emitting call's timer-tick delta, with
presentation-build/projectable-scan counts in value0/value1; it is not a
filesystem copy count. PauseProjectionUs.value1 is the conjunction
rateCleared && noTimerTick success flag. The earlier immutable A-only preparation
misdescribed that field as a tick delta; this combined document supersedes that
description while preserving every raw value. RepeatedPauseProjection.value1
is an actual repeated timer-tick delta.

A used operator V5 and B used V6. The V6 change corrected reporting selection;
the functional build/case/control code was unchanged, with 15 safe selector
controls recorded separately. The original A parent selector/reporting failure
and exit 1 remain preserved. Its later read-only exact-ID reconciliation did not
rerun A or relabel raw failures, canceled runs or preparation attempts. Primary
identity uses the actor-returned exact run ID and PID, contained archive and
report. Capture.StartedUtc is a held invocation/launch-window bound, not an
observation of kernel ProcessStartTime.

Public files omit raw physical paths, titles, process/thread and machine identity
metadata. Originals and their full private path inventory remain in governed
evidence. Minimal source digest/byte/row references bind this deterministic
subset; the observed post-archive v1 summaries are labelled separately and are
not claimed to be archive-inventory sealed. Original per-case statuses and
failure reasons remain available. The archive validator checks packaging, not
the semantic interpretation of these observations.
