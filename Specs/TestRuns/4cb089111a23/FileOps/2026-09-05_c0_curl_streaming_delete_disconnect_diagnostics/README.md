# C0 bounded Curl Delete — expected disconnect diagnosis

Starting checkpoint `5132d866` plus the bounded Delete candidate, exact-root-error
fix and assertion diagnostics. Command:
`Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_ -FailFast -RunId c0-curl-streaming-delete-diagnostics-debug-20260905`.

Build: 2:45, zero warnings/errors. Receipt
`10159736024408df27f546d00c5078bfed075cdea1e0aefc3a220f079ca80bf7`.
Harness: 9 passed, 1 failed, 3 skipped (not reached after failure); native
partial-tree assertions: 34 passed / 3 failed. Both independent selected-item
proof controls now pass. This remains a failed run, not qualification evidence.

All four partial-tree scenarios report E_ACCESSDENIED, exactly two mutations,
the expected remaining files/root, and truthful indeterminate completion. Three
fail only because the fake server records WSAECONNRESET (0x80072746) while the
streaming consumer closes an unfinished parent LIST. The fourth passes depending
on disconnect processing order. The diagnostic metric records each predicate.

The fixture correction explicitly treats this exact peer reset like existing
EOF/cancel disconnects and drains the endpoint before asserting health. It does
not ignore other socket errors or relax any content, request, or receipt assertion.
The next full C0 run must still pass the late-listing and bounded traversal cases.

All results/reasons, traces and the complete ordered perf stream are copied from
the exact run-owned C:\RedSalamander.Perf artifacts. Retained-run disk-audit
limitations remain in the aggregate; no older evidence was removed.
