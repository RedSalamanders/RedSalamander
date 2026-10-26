# Retained multi cache still loses blocking-transfer connections

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_ -FailFast -RunId c0-curl-entry-pool-c0-debug-20260905`.
Build 02:41, zero warnings/errors. Receipt
`cb0eec27e06ec1297fc08a045c2612b8241910adf493b366003c15a6ea6271a4`;
workspace `e2d87c13c40c5dcf4008977aca6a9252cb595b8eedcf2459df5192e02846d2f8`.

15 passed, 1 failed, 1 not reached (FailFast). Copy 139/141, all 46 scenarios
finish. Lookup has its separate same-source 197/197 pass; it is not substituted
into this broad run. Delete 127/127, size 85/85 and preflight 97/97 pass.

Parallel wide Copy publishes all 4,097 files in 114,625,496 us with correct contents
and receipt, but fails the connection-count guard: 4,059 listener creations and
12,332 reuses. Serial Copy then refuses a connection after 468/4,097 publications
in 7,481,941 us (472 creations, 1,407 reuses). Both fixture endpoints remain healthy.
System/Tcpip 4227, record 324993, is observed at 19:36:21 local (UTC+02:00).
The read-only query covered 17:33:20–17:37:05 UTC, provider Tcpip, IDs 4227/4231.
No TCP setting, corpus, deadline or safety proof changed; no unrelated process killed.

The installed header reports libcurl 8.21.0. Its
[matching source](https://github.com/curl/curl/blob/curl-8_21_0/lib/multi.c#L468-L473)
destroys an easy handle's private multi cache when attaching it to a caller-owned
multi handle. Thus alternating cursor LIST and blocking transfer on the same easy
handle still closes transfer connections, even when the cursor multi survives.
Next: separate cursor and blocking easy handles within each existing exclusive pool
entry, reset both callbacks on return, retain the same idle/expiry limits, and test
alternating operations explicitly. No second global pool or new transfer scheduler.

Raw perf: 13,960,661 bytes, 39,952 rows; SHA-256
`abd468f1ea318db7de2c374dbf178c84b5f94e859fc7a3df93b13b57a300fcbc`.
All Curl/C0 rows are retained in original order with `metric-selection.json`.
Results/reasons and both traces remain complete. Disk audit 81 issues; none deleted.
C0, Release and final Fresh Full remain open.
