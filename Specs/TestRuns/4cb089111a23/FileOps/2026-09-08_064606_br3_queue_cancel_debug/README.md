# Queue-paused cancellation regression

Source: `c36d3497a0d5214168c062bf3f67a16e6baa909a` plus the uncommitted
P1/P2/BR-3 patch. Test-enabled x64 Debug, zero build warnings/errors.
Build receipt: `7c419ffc27df7669e7d01415ce5ee4ed9c6f67c25e676b17df9893cde0039417`.
Runner: `20260908T064606Z-66052-62ad35df8dd940b5baf986310f2339ed`.

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -CaseFilter BR3_CancelQueuePausedTransfer -TestRoot Z:\RedSalamander.Perf`.

All three setup/case/cleanup results passed. The metric
`FileOps.SelfTest.QueuePausedCancelLatency` records 62,000 us against the
3,000,000-us bound. Both real Local copies had reached host pause waits before
cancellation; the other transfer remained paused through the completion assertion.
No queue release or predecessor completion qualified the pass. Results, trace and
the complete bounded metric stream are retained verbatim alongside this note.

## Observed RED control

Before adding the cancellation/stop check at the top of `WaitWhilePaused`, the
same instrumented fixture ran as
`20260908T064049Z-59444-dbd85e683bd044da8b0a2f23f3933019` (Debug, zero build
warnings/errors, receipt
`548573d5223184f6cf5c0cc1d55e19329a34ddc10f873ec4579d6c804c937f62`).
The tool-observed result was 2 passed / 1 failed, with the case failure:
"Cancel did not finish the queue-paused transfer within 3 seconds while its
predecessor remained paused." Setup took 1,313 ms, the failing fixture 12,203 ms,
and cleanup 31 ms; process results reported duration 13,547 ms. Failure-only
queue release allowed both tasks to drain.

This paragraph preserves the execution observation, not a reconstructed raw result
or metric. The runner's next-launch stale-run cleanup removed the RED artifacts
before promotion. Therefore this archive's raw performance authority is the
candidate only; it does not establish a numeric before/after speedup. It also does
not close BR-3's whole-process Phase 7 or Fresh Full requirements.
