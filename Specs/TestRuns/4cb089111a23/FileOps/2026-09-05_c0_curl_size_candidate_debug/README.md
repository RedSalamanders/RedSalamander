# C0 iterative Curl size — Debug family candidate

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_ -FailFast -RunId c0-curl-size-iterative-final-debug-20260905`.

Build: 165,960 ms, zero warnings/errors, tests enabled. Receipt
`232a4b12b1c3f0aab920d2862767bac8d1bd2dfb224d9004359bcb4b6ff730db`;
source `0620c05e5bd9dfc926d9014e8b7e2e7001b40caa0576c78e04f24dac7da1abbb`.
Host: 14/14, no skips, 24,094 ms suite duration. Native guards/lost reply/partial
tree/late listing/bounds/size: 29/33/37/49/127/79 passes, zero native failures.

The idle-handle reset check passes. Blocked LIST cancellation returns cooperatively
in 243,352 us. Wide 4,097 / depth-80 size durations are 68,692 / 429,313 us. The
same-machine red exact-case baseline recorded 53,548 / 269,338 us with incorrect
feedback; this candidate also runs after the broader C0 group. These are single
Debug samples, not a throughput improvement claim. Maximum admitted size-walker
reservations: 129 frames, 124,624 path bytes, 6,624,408 metadata bytes; not whole
process memory or IMAP qualification.

Both traces, case results/reasons and aggregate are unchanged SHA-256-verified raw
copies. The oversized perf stream remains in the exact run under
`C:\RedSalamander.Perf`. `metric-selection.json` binds its digest, byte/row counts,
deterministic metric-name filter and original line numbers; all 292 selected Curl/C0
rows are preserved in order. This candidate predates the fixture's witnessed
Winsock-abort correction and the later partial-final-feedback repair. It is not final
repository evidence. Existing disk-audit findings were retained, not cleaned away.
