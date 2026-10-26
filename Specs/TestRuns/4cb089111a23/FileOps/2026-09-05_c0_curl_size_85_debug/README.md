# C0 Curl directory-size repair — final Debug qualification

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_ -FailFast -RunId c0-curl-size-85-debug-20260905`.

Build: 151,403 ms, zero warnings/errors. Receipt
`2788edb5893eb438cdab3dd77e783ffdbb7c78f1cc78b2fae4c48e2d5c05f7b9`;
source `3663194d725a23644d1ff88b3eeb10eb4d3955600ed513571a41e0a791c190df`.
Host: 14/14 passes, zero failures/skips; 24,484 ms suite duration.

The run passes the unchanged native Delete guards/lost reply/partial-tree/late-listing/
bounds gates (29/33/37/49/127 assertions) and the expanded directory-size gate
(85 assertions: 26 scenarios plus ABI, pool-lifetime and startup checks). The two
partial-final-feedback scenarios return cancellation / callback failure rather
than hiding them behind accumulated partial status; known totals are retained.
The earlier malformed-list Winsock-abort failure is preserved separately; its
fixture-only correction does not relax exact contents, status or receipt checks.

Blocked LIST cancellation: 245,950 us, cooperative without emergency teardown.
Wide-4,097 / depth-80: 86,572 / 342,381 us. Maximum admitted reservations across
26 traversal samples: 129 frames, 124,624 path bytes, 6,624,408 metadata bytes.
These reservations exclude opaque transport and transient allocations. IMAP
buffering and real FTP/SSH session limits are not qualified by loopback FTP.

The same-machine red Debug exact-case baseline recorded wide/deep 53,548 / 269,338
us and failed cancellation's two-second bound before emergency shutdown. These
samples run after the broader C0 group; do not claim a throughput improvement
or use three samples as a percentile qualification. Correctness and the explicit
cancellation/storage bounds are the claims supported here.

Case results/reasons, both traces and aggregate are unchanged SHA-256-verified
copies. The oversized original perf stream remains beneath the named owned run
in `C:/RedSalamander.Perf`. `metric-selection.json` binds its byte/row counts,
digest, deterministic filter and original line numbers; all 298 selected Curl/C0
rows are retained in source order. This is not a whole-run performance summary.

The three final runs bind the same source snapshot. Subsequent edits only finalize
spec/evidence reporting. Source/localization checks pass 197/197. The intermediate
shared z.dll Clean warning remains a C6 issue despite these warning-free builds;
its exact build log is archived with that intermediate run. Retained disk-audit
findings were not deleted. This closes the size-repair substep only: C0's other
routes, C1–C6, inherited I17 and exact Fresh Full still prevent moving the plan to Done.
