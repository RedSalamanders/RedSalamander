# Exclusive fake-FTP endpoint ownership — verified, broad failure remains

Base `2283f0cf`. Both listening and bound-only endpoints now require
SO_EXCLUSIVEADDRUSE before loopback port-zero bind. Two negative controls verify
the installed option and rejection of a same-address/port SO_REUSEADDR challenger
(10013 in both builds). No client, machine, retry, deadline or corpus change.
This is a fixture-ownership correction, **not a repair of broad deep Move**.

| Run | Native result | Internal checks / wrapper duration |
|---|---|---|
| lookup Release | 3/0/0 | 273/273; 21.8 s |
| lookup Debug | 3/0/0 | 273/273; 23.5 s |
| broad Release | 15/1/1 | Copy 134/142; 82.9 s |
| CRUD Release | 3/0/0 | 3.4 s |
| CRUD Debug | 3/0/0 | 4.1 s |

Policy: 197/197, no skips, 18.1735044 s. Debug full build: 187622 ms, zero warnings
or errors. Release full build: 258597 ms, one MSB3061 Clean warning for shared
`Plugins/brotlidec.dll`, zero errors; not warning-free qualification. Both bind
snapshot `7ce7d8bc9e5f758454a82af5b6a642a83a6da2857e45d85bf123335b52723cfe`.
Release receipt `d7b48c90c98502b6886baea888122f34b8a9ce7acc05c47238b34ff5efa10820`;
Debug receipt `0b1db7e260d764b9b97de84bcff52788151bde74ecb7f97038ea07cc65fdffbc`.
The tested patch and four exact file hashes are retained before final docs updates.

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -TestRoot C:\RedSalamander.Perf -CaseFilter C0_CurlEntryLookupTruth -RunId c0-curl-exclusive-lookup-release-20260906`.
Other exact commands are in the sibling archives. All raw results, traces and
ordered digest-bound metric projections retain their failed/skipped statuses.
145–149 old-run disk-audit findings remain visible; no evidence was deleted.

Broad Debug was not rerun for this fixture correction; its prior wide timeout is
still open. Next inspect/reclaim completed cursor borrows before child descent,
with nested-cursor reuse and terminal-reply/cancel/fragmentation controls. The
active plan keeps every remaining C0–C6 gate; no Done claim or live-provider write.
