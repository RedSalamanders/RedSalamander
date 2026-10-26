# C0 Release — three passes after fake-server disconnect repair

Command: test-enabled `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Release -CaseFilter C0_ -SelfTestRepeat 3 -FailFast -RunId c0-curl-size-final-release-repeat3-20260905`.

Build: 199,866 ms, one warning, zero errors. Receipt
`3bcd30c1c7c6d4dae13285079e332fab15a2b810fc01495f196b6d9a9c3a6727`;
source `440e31ecaffa7503e9307e19e10e584fc19bf96de5d554ea736a8725b63a9cad`.
MSB3061 could not delete shared `Release/Plugins/z.dll` during Clean. This extends
the existing C6 shared-output ownership issue; the warning is not resolved by a
subsequent passing build and was not suppressed.

All three C0 repetitions pass: 42/42 host entries, no skips. Size passes 79/79 and
Delete bounds 127/127 in each repetition. One `WSAECONNABORTED` observation is
recorded by `FileOps.Curl.FakeFtp.PeerDisconnect` while exact scenario assertions
still pass, providing a witnessed green counterpart to the diagnostic failure.
Blocked size-LIST cancellation: 244,100 / 246,492 / 244,128 us, all cooperative.

Both traces, all case statuses/reasons and aggregate are unchanged SHA-256-verified
copies. `metric-selection.json` binds the oversized original perf stream and all
877 selected Curl/C0 rows, kept in original order. Full raw evidence remains in the
named owned run under `C:\RedSalamander.Perf`; no prior run was removed.

This is intermediate evidence: the subsequently found partial-final-feedback
precedence bug and its two new scenarios are not covered here. Final qualification
must execute the expanded 85-assertion gate. C0 and C6 remain open.
