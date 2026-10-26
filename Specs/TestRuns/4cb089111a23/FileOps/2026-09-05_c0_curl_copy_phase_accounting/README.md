# C0 Curl Copy phase timing and accepted-transfer accounting

Run `c0-curl-copy-phase-accounting-debug-20260905`, Debug x64, based on
`e68a8d0b786fcb4f4ee63c7451fdcc4739f07d3b` plus the two Curl instrumentation/test
files and plan/performance-spec changes. Source snapshot:
`f915a122a1f7811e1f1805a8f3a88b60d7d07f5f1b6a87c72695712839d3059d`.

**Still red: serial wide Copy exceeds its deadline.** Native and governed results:
15 passed, 1 failed, 1 not reached; runner exit 1. Copy has 141/142 assertions,
all 46 scenarios finished. Delete 127/127, size 85/85, Move preflight 97/97 pass.
The later lookup case is not reached even with `fail_fast=false`; its modified
accepted-transfer checks still need an exact-case run. Do not claim new 201/201 evidence.

## Measured result

| Wide Copy | Parallel (4) | Serial (1) |
|---|---:|---:|
| Published / requested files | 4,097 / 4,097 | 3,210 / 4,097 |
| Elapsed microseconds | 115,640,339 | 120,008,558 |
| Native HRESULT | S_OK | 0x800705B4 (timeout) |
| Exact complete contents | yes | no; partial result |
| Truthful completion | yes | yes |
| Source listener creations / reuses | 5 / 4,094 | 2 / 3,211 |
| Destination listener creations / reuses | 8 / 12,284 | 2 / 9,632 |
| Destination LIST commands | 8,195 | 6,423 |
| Destination generated LIST bytes | 873,274,394 | 536,285,269 |
| Destination LIST generation + mutex wait, us | 106,558,912 | 49,300,948 |
| Sum of reserve-staging-name phases, us | 182,489,925 | 52,261,309 |
| Sum of promote phases, us | 177,374,090 | 54,071,158 |

Parallel accepted LIST/RETR/STOR first-use/reuse totals match every command and
listener. Source RETR and destination STOR each have 4 first uses + 4,093 reuses.
There is no NLST or abandoned passive request. The previous >=4,096-reuse oracle
incorrectly assumed one session; complete accounting now passes without relaxing
the 64-creation ceiling, 4,097-file corpus, exact bytes, or 120-second deadline.
Serial accounting remains red because it does not complete all payload transfers,
not because of connection churn. The refused-RETR lifecycle control passes its
new assertion: the rejected request is not counted as an accepted transfer.

Phase rows are summed file wall time, not CPU time; parallel values overlap.
Server generation time is contained within client waiting time and must not be
added to it. These measurements locate expensive repeated destination validation,
but do not separate production parsing from all transport/scheduling costs.
The fake server's serialization/mutex cost is substantial. No throughput repair
or real-server performance improvement is claimed by this instrumentation change.

Same-machine [previous baseline](../2026-09-05_c0_curl_split_handles_wide_red/README.md):
parallel 117,448,331 us with 4,097 publications; serial timeout with 3,309.
Single-run timing/partial-count variation does not establish an improvement or
attribute a regression to instrumentation. Next: make one measured change while
preserving strict all-row/late-error/collision validation, then rerun the same corpus,
lookup controls, R0f CRUD regression, and test-enabled Release.

## Provenance and reproduction

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_ -RunId c0-curl-copy-phase-accounting-debug-20260905
```

Use a fresh ID for another run. Build 03:00, zero warnings/errors; receipt
`9364f92977cdb862a4c835ee1576abce9e7179e39751d94fcc87307009bba4c3` is
application-target evidence, **not** a full-solution receipt usable with `-SkipBuild`.
Entry fingerprints, complete native/root/runner results, traces and build log are
retained unchanged. Suite 305,375 ms; governed wall 5m12.2s. Disk audit: 88 retained
issues, none manually removed. C0 remains VERIFY; this is not final qualification.

Raw perf remains under this run in `C:\RedSalamander.Perf`: 20,434,070 bytes,
58,495 rows, SHA-256
`3ba5b5859b635a2ead96b49805cffda62659eb12a2ae39fe161f21da153f2e77`.
The archive retains all 1,191 `FileOps.Curl.*` / `fileops.c0.*` rows in original
order. `metric-selection.json` binds every original one-based line and both digests;
the projection is independently verified against the raw stream before acceptance.

Post-run source/resource policy: 197/197 passed, zero skips, 12.750 s;
`source-resource-contracts.xml` is the unchanged NUnit result from
`c0-copy-accounting-host-policy-20260905`. Spec inventory: zero blocking findings.
These checks do not replace the unexecuted lookup runtime checks or repair Copy.
