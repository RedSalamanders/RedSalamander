# Instrumented ASan failure diagnostics

Source `b60d16c5a1cea3efc39cf04f85c5ac9a95be4fbc`, x64 ASan Debug, built with zero
warnings/errors; receipt `c5c2edec87ac4bc5339f004aee2879fa05369b7bddb16ab4cabd4ee2b8bd4e29`.
Original predicates and budgets were retained. These failed diagnostics are not
repository qualification and remain immutable baseline evidence.

The owned initializing-progress gate held the volume mutex. The optional status
snapshot returned only after 5,000,855 us, with its gate timed out, against the
500 ms requirement. The native result is one failed case. Its wrapper also failed
coverage because the new name was omitted from the explicit native case list;
both complete results are retained. That registration omission must be repaired,
without ignoring the coverage failure or promoting this wrapper result.

The UI predecessor cluster recorded 10 passes and three failures. During the
unchanged 1,000-command retention measurement, private bytes grew from 692,473,856
to 1,180,155,904 while live ASan bytes fell from 57,853,237 to 57,278,958. Allocator
heap bytes grew from 270,794,752 to 697,626,624, and free bytes from 105,490,736 to
231,730,268. These distinguish cold instrumented allocator growth from live
retention; they are not a production memory improvement or general no-leak claim.
The measured private-byte assertion still failed and was not promoted.

Breadcrumb Dismiss remained absent after a three-second post-failure observation;
the card became compact. The first snapshot was taken before terminal auto-collapse.
The inspection-return failure reproduced with correct native/logical focus on A
but empty default/cancel bindings. Its predecessor cluster retained all predicates.
The final close case passed at 84,336 us against its unchanged 100 ms gate; model
retirement and deferred-disposal handoff took 179 and 58 us. The earlier Full's
112,377 us overrun remains historical; its exact cause is not established here.

The selection replay recorded three target passes and two failures, plus ten
passing setup/cleanup cases. In its two failed Auto-4 leaf scenarios the original
partial-byte sampler reported two and zero, while explicit admitted transfer
phases recorded four both times. All thirty shape/configuration observations
retained the expected real phase concurrency; content/counting checks were intact.
This does not justify substituting configured limits for actual work.

Complete native/wrapper verdicts and digest-bound selected metric/trace rows are
listed in `manifest.json`. Original streams remain in the named local completed
runs. Production corrections, revised fixture qualification, six builds and final
Fresh Full runs remain required. No sanitizer option was changed or allocator purged.
