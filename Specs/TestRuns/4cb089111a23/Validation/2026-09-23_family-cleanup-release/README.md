# File Operations failed-family cleanup qualification

The Phase5 discovery family passes three consecutive runs after correcting its Start-now assertion
and draining outstanding operations before family cleanup. The added regression leaves a throttled
copy paused and another copy queued, calls the same cleanup drain, verifies both cancelled completion
receipts, and proves a fresh copy completes intact after their locks are released.

- Source: isolated `C:/RsI26Final`, detached `ccebd5af68eb45f2e9deed64a7063780e71eb1ce`
  plus the 13 explicit inputs in `raw/candidate-inputs.json`. DxUi pin remains `78b3de389a189c7f86f611787e0489fb6d474218`.
- x64 Release test-enabled build: zero warnings/errors, receipt
  `0b0de26d09d4fb16e2b0bc7411a119e90ec23cb4aa40d7d9d160ef9bf629275f`.
- Run `20260923T112251Z-68784-055a5dd6113545268624965817f2bad4`: 33 passed, zero failed/skipped,
  54.6 seconds, sandbox audit zero. No quarantine or reused outcomes. Nonactivating harness only.
- Raw trace contains three successful paused/queued cancellation drains. Start-now checks only the
  selected scheduling policy; the worker waiting indicator may remain set for a mutation interlock.
  Other queued tasks/global mode and eventual operation entry/completion remain asserted.

The earlier canonical FileOps run `20260923T105425Z-9612-799a01e299d5408e92a617fa88237d83`
failed the immediate waiting-indicator assertion. Its cleanup did not cancel the throttled operation;
four subsequent families timed out with that operation still retaining an interlock. The verified
owned test PID 36076 (runner parent 9612, exact self-test command/path) was stopped during family 19.
That run remains incomplete and failed, including its result-coverage failure. The ownership receipt
and complete trace are retained here. Full original runner files and its 385,466,119-byte metric
stream remain under `C:/RedSalamander.Perf/evidence/i26-ui/footer-summary-final-20260923/fileops-runner-artifacts`.

This packet qualifies the focused family only. It does not establish Full, the final UI path/menu
candidate, native focus/DPI192/AT, combined library-pin or resource acceptance. A later UI candidate
also changes ordinary path abbreviation and destination-menu presentation and needs its own receipt.

`SHA256SUMS.txt` binds every byte-preserved raw payload. Full focused originals are retained under
`C:/RedSalamander.Perf/evidence/i26-ui/family-cleanup-20260923`.
