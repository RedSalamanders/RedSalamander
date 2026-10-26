# C0 native Curl Delete late-listing ordering — red baseline

- Starting checkpoint: `cbf19004`, plus only the new fixture, host test dispatch,
  and validation/plan documentation. Production walkers were unchanged.
- Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_CurlNativeDeleteLateListing -FailFast -RunId c0-curl-late-listing-red-debug-20260905`.
- Debug build: 2:44, zero warnings/errors, tests enabled. Receipt
  `f588db790ceb90110d5f1f386e7a272dc5709d66ff9ce01aed95ec92ae7178f8`;
  source snapshot `1ec4af59a32253bf0f315260377248511623af4bb75d3e3cf0f432f4a9524bd7`.
- Harness: two setup/cleanup passes, one failed behavior case, zero skips.
  Native assertions: 31 passed, 18 failed across twelve scenarios.

All six concurrency-1 scenarios pass: single/bulk Delete each completes ordinary
work, reports a truthful partial result after a late refusal, and cancels a late
blocked LIST after deleting the earlier leaf. Cancel acknowledgement takes
1,001,326 / 1,004,766 us (single/bulk), with one deleted leaf and cooperative return.

All six concurrency-4 scenarios fail the ordering witness. The server waits for an
exact earlier DELE before returning `/selected/b`'s listing; the provider's eager
collector sends no earlier deletion and hits the fixture's five-second gate timeout.
Observed durations are 5,008,230–5,016,793 us. The gate never holds the fake server's
command mutex, so it does not prevent other sessions from making progress.

Each parallel scenario fails three related expectations: ordering, the intended
post-deletion state, and its intended terminal receipt. These are six pre-scan
failures, not eighteen independent production defects. In particular, an untouched
tree's no-commit receipt does not itself violate the existing proof contract; the
test requires reaching the deliberately post-mutation failure/cancel scenario.
The 723/743-us parallel cancel-return samples are not successful cancellation
qualification: their post-deletion gate was never reached.

This small-tree baseline demonstrates the eager-discovery dependency only. It does
not prove wide/deep bounded memory, retired recursion, UI responsiveness, live SSH,
or Done readiness. Next replace the collector with bounded iterative streaming,
extend deep/wide/queue evidence, and rerun this witness plus existing C0 regressions.

Raw results, aggregate, main and FileOps traces, and complete perf stream are
preserved from the run-owned `C:\RedSalamander.Perf` artifacts. The runner records
43 retained-run disk-audit issues; no prior evidence was removed. Source/localization
checks pass 197/197, and the spec inventory has zero blockers. These checks do not
turn the deliberately failing runtime baseline into a passing qualification.
