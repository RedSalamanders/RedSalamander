# Bounded lookup — broad regression exposes cursor connection churn

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_ -FailFast -RunId c0-curl-entry-lookup-bounded-c0-debug-20260905`.
Build 02:53, zero warnings/errors. Receipt:
`9b6f0acf03207c98f253dd3ec79be59d47935235b666047987d0ce83dd46da18`.
Workspace `0439d4b0f9cde5bac34ac0bfce9bdabd7fe82bddb64e7ec517d7dad10f887c66`,
identical to the [197/197 focused pass](../2026-09-05_c0_curl_entry_lookup_bounded/README.md).

Result: **15 passed, 1 failed, 1 skipped**. Thirteen named C0 predecessors plus
Setup/Cleanup pass. Copy is 123/141; the new lookup case is not reached because
FailFast stops after Copy. Its separate focused pass is not substituted into this run.
Delete, size (85 checks), and Move preflight (97 checks) remain green.

Wide parallel: 1,239/4,097 publications in 13,837,671 us, 3,689 listener creations
and 1,271 reuses. Serial: 143 publications in 2,069,301 us, 435 creations and 143
reuses. Both fail ERROR_CONNECTION_REFUSED with healthy fixture endpoints. Later
deep/depth-limit scenarios also fail to connect. No deadline/corpus/proof was relaxed.
Windows System/Tcpip event 4227, record 324992, at
`2026-09-05T17:19:51.0140914Z` confirms outgoing endpoint reuse exhaustion within
the test window. The bounded read-only query used ProviderName Tcpip, IDs 4227/4231,
17:18:40–17:20:10 UTC with ErrorAction Stop. No TCP settings or unrelated processes changed.

Cause: every new lookup opens/destroys its own cursor multi handle. Destruction
closes that handle's connections, even when the easy handle returns to its pool.
This follows [libcurl's multi cleanup contract](https://curl.se/libcurl/c/curl_multi_cleanup.html)
and the measured listener counts. Next repair: retain the worker-exclusively borrowed
multi connection cache with the existing bounded easy-handle pool, while detaching
the easy handle and clearing borrowed callbacks at every return. Keep pool expiry,
shutdown/quiet-point rules and strict listing validation; do not add a process-global
unbounded cache or revert to whole-directory buffering.

Raw perf: 5,193,982 bytes, 14,804 rows; SHA-256
`46c9d8e68c7dbeab62dfbb96e84fe1dc4409d6bcd5cc88483194f6ca473a943c`.
All 659 Curl/C0 rows remain in original order, bound by `metric-selection.json`.
Complete results/reasons and both traces are retained. Disk audit: 79 issues,
none deleted. C0, test-enabled Release, ordinary FTP and final Fresh Full remain open.
