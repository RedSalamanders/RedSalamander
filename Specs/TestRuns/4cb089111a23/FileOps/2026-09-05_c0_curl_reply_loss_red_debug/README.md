# C0 Curl lost mutation reply — pre-fix Debug witness

- Source: `58f22d0e` plus prior C0 changes and the new fault fixture; production
  `CurlPerformQuote` still used its three-attempt loop. Same machine `4cb089111a23`.
- Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_CurlCommittedMutationResponseLost -FailFast -RunId c0-curl-reply-loss-red-debug-20260905`.
- Accepted build receipt: `4d8b2d199891f417ada745a1b4a8caa7de23a067cb22f148d982de6805a85d6e`;
  0 warnings/errors; build 3:01.
- Harness: 2 passed, 1 failed, 0 skipped. Native checks: 21 passed, 12 failed.
  This is a reproduced product defect, not an unresolved fixture prerequisite.

All four normal single/bulk Delete/Rename controls passed. Each lost-reply case
performed the original mutation, atomically planted different bytes at the source,
and dropped the control connection before acknowledgement. Each then sent the
mutation twice: Delete removed the replacement; Rename overwrote the first result
with the replacement. All four also reported successful completion. The native
state, request-count, and terminal-truth assertions each failed (12 failures).

`FileOps.Curl.MutationReplyLoss` preserves all eight raw rows. Lost-reply durations
were 219,580 / 219,870 / 219,387 / 221,295 us for DeleteItem / DeleteItems /
RenameItem / RenameItems. These measure the unsafe replay path, not useful throughput.
The later candidate adds an explicit no-incidental-NLST assertion and host Delete
feedback coverage; neither extra assertion is claimed as exercised in this red run.

The four evidence files are intact allow-listed copies from
`C:\RedSalamander.Perf\runs\c0-curl-reply-loss-red-debug-20260905\artifacts\selftest\last_run`.
The runner reported 33 retained-run disk-audit issues. No old runtime data was
removed; this is not clean-root or final Fresh Full evidence.
