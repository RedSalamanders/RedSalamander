# Throughput UI — unchanged-production witness

Run `20260919T140946Z-30692-40a1c88c38c146e5b487e81e1b92807e`, x64 Debug at HEAD
`0a926ea562b3c671f3f3d52c1af527c8c49b2bd0` plus the fixture/debug enumeration additions.
Receipt `1c2244cbc23c4dddaef1c5a8ccc8279f7994f9bb3953dc5d9fdc947183e44013` binds the inputs.
The build has two expected unhandled-new-enum warnings and zero errors. The case fails 0/1 at
the missing disclosure assertion. Disk audit reports one pre-existing direct-launch scratch folder;
that folder was later preserved under the governed evidence directory. This is not a green run.

Five ordered eight-redraw expanded batches: 373861, 374500, 374879, 373811, 376804 us;
median 374500 us. The full metric stream, result and trace are retained. The measured fixture is
unchanged in the [first candidate](../2026-09-19_161917-i26-throughput-ui-first/README.md).
Original fixture source remains at
`C:/RedSalamander.Perf/evidence/i26-ui/throughput-ui-baseline/fixture-baseline.cpp`.

Archived runner output has trailing whitespace normalized; unchanged full logs remain in the governed local evidence folder.
