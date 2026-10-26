# C0 Curl directory-size — first iterative candidate

- Starting commit: `1337dee1`; uncommitted iterative sizing/callback repair.
- Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_CurlDirectorySizeTruth -FailFast -RunId c0-curl-size-iterative-candidate-debug-20260905`.
- Build: 193,004 ms, zero warnings/errors, tests enabled. Receipt
  `bdf4e28e4ba5a45608c25d75c5ddd92531d7c1fc70e2b5b57cf1d921e9ab5ec5`;
  source `6f2f39ad8778019c410f11ed1a3021887766682275bd9377967a00aa9d2883a1`.
- Host: 3/3 passes, no skips; native: 77/77 assertions over 24 scenarios plus
  four argument checks. All scenarios preserve server contents and honest feedback.
- Blocked LIST cancellation: 241,591 us, cooperative return, no emergency stop.
  The same-machine red baseline needed emergency teardown after failing its
  two-second cancellation bound (overall call 7,103,065 us).
- Wide 4,097 / depth-80 timings: 61,950 / 383,315 us versus red 53,548 / 269,338 us
  (approximately 16% / 42% slower). These single Debug samples are not a throughput
  qualification; the new checks and open ancestry cursors have costs. Real FTP/SSH
  session limits and latency remain open and must not be hidden by correctness.
- Maximum admitted traversal reservations: 129 frames, 124,624 path bytes,
  6,624,408 metadata bytes. These exclude opaque transport/transient allocations
  and IMAP's buffered adapter; they are not whole-process memory measurements.

The five raw files are exact SHA-256-verified copies from this owned run under
`C:\RedSalamander.Perf`; the entire 144,358-byte perf stream is retained. Disk audit:
52 retained-run issues, none removed. This run predates the return-to-pool option
reset, final-feedback latch, and two pool lifetime assertions. It is intermediate
evidence, not the final candidate or repository-wide closeout. C0 remains VERIFY.
