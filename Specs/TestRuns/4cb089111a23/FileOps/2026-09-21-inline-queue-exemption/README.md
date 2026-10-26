# Inline F2: retain the global Queue exception

x64 Debug, machine `4cb089111a23`, base `20ed39ef` plus the bound dirty source snapshots.
DxUi stays pinned at `78b3de389a189c7f86f611787e0489fb6d474218`.

The existing interlock fixture now reapplies Queue with an explicit user pause and
while a second inline rename waits on the real overlapping-root interlock. Before
the scheduler change it fails in 188 ms because the waiter acquires global Queue
waiting. The repair excludes typed inline renames from live Queue waiting and from
the oldest-active selection used to calculate queue-induced pauses. Explicit user
pause, mutation interlocks and cancellation retain their existing authority.

| Run | Passed / failed / skipped | Result |
|---|---|---|
| baseline | 2 / 1 / 0 | Meaningful regression, including setup/cleanup |
| candidate | 9 / 0 / 0 | Three complete rename repetitions, 703 / 625 / 734 ms |
| phase05 | 10 / 0 / 0 | Discovery, queue ordering, Start now and cancellation |
| r4a19-initial | 4 / 1 / 0 | Provider control passes; alternate-volume flag omitted |
| r4a19-qualified | 5 / 0 / 0 | Corrected invocation, including independent volumes |
| broad-initial | 2 / 1 / 210 | Stops on copy-conflict layout assertion; no broad pass |

Both full-solution test-enabled Debug builds pass with zero warnings/errors.
Baseline receipt: `f29e8a107693dc831baa0cf47f221a9a40a2b84a03cb85b70b8aae4c852d2edc`.
Candidate receipt: `3ce4eb37ee840e24ad7dd1653902f1b0c2c9118e406e9e92d13d9f8afad3f092`.
Full build/runner logs and native family archives are retained under
`C:/RedSalamander.Perf/evidence/i26-ui/inline-queue-*20260921`.

The historical Full timeout witness contains 182634005 us of pause after the second
rename acquired its actual interlock. The selected raw rows and their manifest bind
the original bytes and row count, preserving row order. The strengthened fixture
fails earlier by assertion; comparing its failed duration with successful complete
rounds would not establish a performance improvement. No FPS, memory or broad
resource non-regression is claimed. Routine runner cleanup removed the unpromoted
exact-run metric streams; their preserved results, traces and aggregates remain.
The older timeout stream and subsequent family metric streams were copied to the
external evidence root before cleanup. Disk audit findings remain explicit (260).

The broad copy-merge assertion reports `promptBeforeActions=false` at 203 ms. A
separate exact classification reproduces it at 156 ms. Its snapshot currently
compares unbounded scroll-content coordinates with fixed action coordinates;
inspect visible clipping before changing product layout or assertions. This remains
an open gate. The earlier Full suite and I26 remain incomplete, despite the focused
Queue and discovery passes. No user or other-owner performance waiver is inferred.