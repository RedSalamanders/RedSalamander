# C0 native Curl Delete guards — pre-fix failure

- Source: `58f22d0e` plus prior receipt repairs and the new native fixture; no Curl
  production repair yet. Build receipt:
  `e4feaff9c676cffd94ebf3af821d44ef36db1ee160990b64aff2799108a93b17`.
- Command: `Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_CurlNativeDeleteGuards -FailFast -RunId c0-curl-native-delete-red-20260905`.
- Build: zero warnings/errors. Harness: 2 passed, 1 failed, zero skipped.
  Native fixture: 14 checks passed, 7 failed. This archive deliberately preserves RED evidence.
- Raw results, trace, aggregate, and complete 116,859-byte metric stream are retained
  unchanged from the run-owned `artifacts/selftest/last_run` directory.

Eight `FileOps.Curl.NativeDelete.RootGuard` rows cover singular/bulk Delete,
concurrency 1/4, and recursive/nonrecursive flags. Serial recursive singular and
bulk calls each sent four destructive commands and left zero files (two were seeded).
Four nonrecursive root calls each sent RMD instead of refusing locally. Parallel
recursive root calls preserved the fixture and returned the expected refusal.

Two `FileOps.Curl.NativeDelete.ListCancel` rows cover the first stalled LIST.
Single-item Delete did not return in its 2,515 ms observation window and required
fixture-owned unblock; bulk Delete returned cooperatively in 1,000 ms. Neither sent
DELE/RMD. These are loopback fault witnesses, not live-endpoint or throughput proof.

The runner disk audit reports 28 retained older-run issues. No older data was
deleted; this archive does not claim a clean sandbox inventory or Full qualification.
