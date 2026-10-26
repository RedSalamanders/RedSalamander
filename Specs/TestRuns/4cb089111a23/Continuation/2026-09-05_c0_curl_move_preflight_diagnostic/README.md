# C0 native Move preflight — diagnostic red baseline

Starting production implementation: `75b20b3e`, unchanged. The fixture puts depth
refusal last and emits a scenario-begin metric. Command:
`Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_CurlMovePreflightTruth -FailFast -RunId c0-curl-move-preflight-diagnostic-debug-20260905`.

Build: 3:05, zero warnings/errors; receipt
`c8052f3072a18668267f62551f3ac767f45e06713ca86f7b6d90dcdc91649a8f`;
snapshot `f5598dbcd395465487800ed7f20bd2d98f167640d96da9d2e10e65da0d529759`.
PID 48140 exits without case results after 19.3 seconds. The final metric explicitly
begins `shape=2;MoveItem;concurrency=1` (depth 129), with no terminal row. The earlier
unclean run is separately retained. No usable stack/exception code is available;
the cause of process termination is not claimed as proven.

Twenty-eight completed scenario rows precede that termination:

- Eight shallow/depth-80 single/bulk controls succeed with exact contents/receipts.
- Four malformed and four overlong listings fail with file-not-found instead of
  invalid-data / buffer-overflow; contents remain unchanged.
- **All four out-of-root listing cases issue four mutation commands** and violate
  the no-premature-mutation contents/receipt assertions. Singular status is
  `0x8007003A`; bulk status is partial-copy, not invalid-name. This is a concrete
  safety failure, not only inaccurate diagnostic wording.
- Four wide unknown-tail controls reach all 4,097 SIZE probes and reject before
  mutation, in 944,498–993,523 us. Do not introduce an aggregate queue-size cutoff.
- Four blocked LIST cancellation controls already pass, in approximately 1 second;
  the owned emergency unblock is not used.

The incomplete process cannot produce a legitimate 97-check suite verdict. Do not
invent case results or classify unexecuted depth variants as passes. Replace unsafe
preflight discovery, then require every scenario to complete and the broad C0 /
ordinary FTP regressions to pass. Native Copy discovery and commitment-map storage
remain separate open work.

The four raw files are exact SHA-256-verified copies from the initialized
`C:/RedSalamander.Perf` run. Raw failures and retained-run disk-audit findings remain;
no old runs, user data, or test payloads were removed/promoted.
