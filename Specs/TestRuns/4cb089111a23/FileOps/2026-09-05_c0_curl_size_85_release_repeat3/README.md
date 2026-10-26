# C0 Curl directory-size repair — final Release repetitions

Command: test-enabled `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Release -CaseFilter C0_ -SelfTestRepeat 3 -FailFast -RunId c0-curl-size-85-release-repeat3-20260905`.

Build: 225,102 ms, zero warnings/errors. Receipt
`4595cc858e3cd5947ccee7d577b129954f4e38010799171794fd70c3ba3ac873`;
source `3663194d725a23644d1ff88b3eeb10eb4d3955600ed513571a41e0a791c190df`.
Host: 42/42 passes, zero failures/skips; 58,672 ms suite duration.

All three repetitions passes the unchanged native Delete guards/lost reply/partial-tree/late-listing/
bounds gates (29/33/37/49/127 assertions) and the expanded directory-size gate
(85 assertions: 26 scenarios plus ABI, pool-lifetime and startup checks). The two
partial-final-feedback scenarios return cancellation / callback failure rather
than hiding them behind accumulated partial status; known totals are retained.
The earlier malformed-list Winsock-abort failure is preserved separately; its
fixture-only correction does not relax exact contents, status or receipt checks.

Blocked LIST cancellation: 242,308 / 243,599 / 250,256 us. Wide-4,097 timings:
7,924 / 10,219 / 9,327 us; depth-80: 235,077 / 274,713 / 278,803 us.
Maximum admitted reservations across 78 traversal samples: 129 frames, 124,624
path bytes, 6,621,312 metadata bytes.
These reservations exclude opaque transport and transient allocations. IMAP
buffering and real FTP/SSH session limits are not qualified by loopback FTP.

The same-machine red Debug exact-case baseline recorded wide/deep 53,548 / 269,338
us and failed cancellation's two-second bound before emergency shutdown. These
samples also differ in build profile; do not claim a throughput improvement
or use three samples as a percentile qualification. Correctness and the explicit
cancellation/storage bounds are the claims supported here.

Case results/reasons, both traces and aggregate are unchanged SHA-256-verified
copies. The oversized original perf stream remains beneath the named owned run
in `C:/RedSalamander.Perf`. `metric-selection.json` binds its byte/row counts,
digest, deterministic filter and original line numbers; all 894 selected Curl/C0
rows are retained in source order. This is not a whole-run performance summary.

The three final runs bind the same source snapshot. Subsequent edits only finalize
spec/evidence reporting. Source/localization checks pass 197/197. The intermediate
shared z.dll Clean warning remains a C6 issue despite these warning-free builds;
its exact build log is archived with that intermediate run. Retained disk-audit
findings were not deleted. This closes the size-repair substep only: C0's other
routes, C1–C6, inherited I17 and exact Fresh Full still prevent moving the plan to Done.
