# C0 Release diagnostic — witnessed Winsock abort classification

Command: test-enabled `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Release -CaseFilter C0_CurlDeleteTraversalBounds -SelfTestRepeat 10 -FailFast -RunId c0-curl-delete-bounds-release-diagnostics-20260905`.

Build: 201,912 ms, zero warnings/errors. Receipt
`f054035c91f7edf1f3539eb2b6f64e91d281daf2b27ad05e5914b02212860dbc`;
source `ebb84582b8a5da67073818f7726e7679c2d2759d66f125a964da6bf89b2ab254`.
Only diagnostics changed from the preceding candidate; assertions were unchanged.

The first repetition fails: two setup/cleanup passes, one failed behavior case,
27 explicitly skipped later entries. This is not ten completed repetitions.
Malformed LIST, single Delete, concurrency four reports actual/expected operation
`0x8007000D`, `contents=true`, `truthful=true`, but server `0x80072745`
(`WSAECONNABORTED`, 10053). The failure is the fake-server health classification,
not lost source/sibling data, a wrong operation verdict, or fabricated mutation proof.

Winsock distinguishes local connection abort from peer reset; see
[Microsoft's error definitions](https://learn.microsoft.com/en-us/windows/win32/winsock/windows-sockets-error-codes-2).
In this witnessed loopback teardown, both terminate the refused listing session.
The repair accepts that narrow disconnected-session observation while retaining
all exact state/status/receipt assertions and recording abort observations separately.

All five raw files are unchanged SHA-256-verified copies from the named owned run,
including the complete 156,740-byte perf stream. Aggregate failure and forbidden
skip classifications are retained. Disk audit: 55 retained issues, none removed.
