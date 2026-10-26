# Split handles stop connection churn; wide throughput remains red

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_ -FailFast -RunId c0-curl-split-handles-c0-debug-20260905`.
Build 02:47, zero warnings/errors. Receipt
`74e5722575c78a670b8e02a3b91e8094f18fd292624ad2693f3949f2b994ed52`;
workspace `00de4818d995672d276fa3a3a66458c7735d2159074102bd2484bd0f96d8071a`.

**15 passed, 1 failed, 1 not reached; Copy 139/141.** All 46 scenarios finish.
The 201-check lookup case is skipped by FailFast; its separate focused pass does
not replace it. Delete 127/127, size 85/85 and Move preflight 97/97 remain green.

Parallel wide Copy publishes all 4,097 files with exact contents and truthful
completion in 117,448,331 us. Aggregate passive listener counts are 13 creations /
16,378 reuses, down from 4,059 / 12,332 in the previous candidate. The fixture still
fails its per-endpoint minimum-reuse assertion. With several healthy sessions,
4,097 transfers do not necessarily imply 4,096 reuses; the aggregate metrics do
not identify the deficient endpoint, so add per-endpoint accounting before fixing
that oracle. Retain the independent 64-creation cap and exact content checks.

Serial wide Copy times out at the unchanged 120-second deadline after 3,309/4,097
publications (119,991,969 us). It uses only 4 listener creations / 13,239 reuses;
both fixture endpoints stay healthy. This is a real throughput failure, not fixed
by correcting the reuse assertion. No corpus, deadline, source proof, or mutation
truth requirement was relaxed. Broad Copy/Release qualification remains open.

Next: account for each endpoint's actual LIST/RETR/STOR attempts and first uses,
then measure repeated full-parent lookup/parsing cost without restoring unbounded
directory caches. Preserve full listing validation and exact requested metadata.
Keep the 201 lookup, 141 Copy and ordinary FTP regression contracts; require a
passing unchanged 4,097-file serial/parallel run and test-enabled Release repeats.

Raw perf: 20,676,317 bytes, 59,433 rows, SHA-256
`be99dbdb1ccfad8d25b0a2fcb19817adfe22953bcb53bba69b67b003fd952227`.
All 681 Curl/C0 rows remain in original order, bound by `metric-selection.json`.
All results/reasons and both traces are retained. Disk audit 83 issues; none deleted.
