# Reusable passive listeners — production wide-copy deadline remains red

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_ -FailFast -RunId c0-curl-passive-lifecycle-repaired-debug-20260905`.

Based on metadata-output repair `61815385`. This revision changes the fake FTP
endpoint, its regression checks/registration and the testing contract, not the
production transport or Copy implementation. Source identities:

- `Shared.cpp`: `3b6239b09730fbc0de8d2cbe22c3e38961b45ff6e18e9df85f62dae4985967b7`.
- `CopyMove.cpp`: `2b6ccce1a2361d89892b24f28dc8ad38ece780606c9100988b4fa68c027cff2d`.

Build: **02:54, zero warnings/errors**.
Receipt: `ab3bedcc8717fc3409b997f3c7f53c5f2364adc1d1ea8fffc6bd91f38e61cd7c`.
Workspace: `685cd7bebb72f4be804e5e7287bb867fcd286d6de986503cfa1c252ca54b6805`.
The preceding [compile failure](../../Continuation/2026-09-05_c0_curl_passive_build_red/README.md)
is retained; no tests ran from that failed build. The governed runner required
process-inventory access; ownership safeguards remained enabled.

Result: **15 passed / 1 failed / 0 skipped**: 13 named C0 passes, one named C0
failure, and successful Setup/Cleanup. Every earlier C0 case passes. Earlier
"15 C0 passes" shorthand included these two harness bookends, not 15 named cases.
The expanded Copy gate is **139/141 assertions**, with **44/46 native scenarios
passing** and both new raw-client lifecycle checks passing. Native suite result:
318,735 ms; Copy case about 268.8 s; wrapper wall time 326,958 ms.

The fixture reuses a consumed session listener but accepts a new data socket for
each transfer. An unused advertisement is retired before the next EPSV/PASV;
the raw-client test leaves a connection queued behind a refused RETR and verifies
two exact subsequent transfers, two listener creations, one reuse and one
retirement. ABOR clears both socket and pending state. Empty idle control sessions
close with 421; partial-command/active-data failures are not suppressed.

| Wide mode | Published / requested | Duration (us) | Listener creations / reuses | Result |
|---|---:|---:|---:|---|
| Parallel, 4 | 3,664 / 4,097 | 120,107,166 | 10 / 14,665 | ERROR_TIMEOUT, promote |
| Serial, 1 | 3,005 / 4,097 | 120,008,132 | 4 / 12,023 | ERROR_TIMEOUT, promote |

Both endpoints are healthy in both rows, with one normal idle closure each and
zero abandoned advertisements. No System/Tcpip 4227/4231 events were present in
the bounded 16:32:10–16:37:45 UTC query (an explicit NoMatchingEventsFound result,
not a suppressed access error). This separates the remaining deadline failure
from the [earlier endpoint-reuse failure](../2026-09-05_c0_curl_copy_metadata_sink_debug/README.md).
The `.PassiveListeners` row remains failed because its full-corpus reuse minimum
is not met by an incomplete copy; no assertion was weakened to call the gate green.

All following depth-128 Copy/Move routes and depth-129 refusals now pass again.
LIST cancellation is 51,099–65,493 us. Exact bytes, outside siblings, early/late
refusal and truthful completion remain required. The ordinary FTP regression
[passes 3/3 on the same snapshot](../2026-09-05_c0_curl_passive_lifecycle_ftp_regression_debug/README.md).
No timeout, corpus size, production publication/size proof, TCP setting or user
data was changed to obtain these results.

## Next production step

1. Capture a small deterministic lookup baseline: malformed/ambiguous parent rows,
   literal and missing names, blocked-list cancellation, and repeated lookups in a
   wide parent. Cover single and bulk Copy/Move/Delete/Rename admission, not only
   the helper. Preserve exact existing-destination and source-retention checks.
2. Reuse `CurlDirectoryCursor` for native entry lookup and remove/replace the
   unbounded `DirectoryEntryCache` path. `GetEntryInfo` in `Imap.cpp` currently
   buffers a whole parent; `DirectoryEntryCache` in `CopyMove.cpp` independently
   caches complete parents and bypasses it. Directory-size root lookup already
   contains a strict cursor implementation that should be shared, not copied.
3. Measure repeated publication lookup work in `GenerateRemoteSiblingPath` and
   `PrepareOverwriteTargetForRename`. Consider avoiding timestamp conversion for
   unrelated entries while preserving full metadata for the requested occupant;
   `ValidateCurlReplaceOccupant` depends on that timestamp. This is an untested
   optimization direction, not a proven speedup or permission to omit collision,
   complete-listing, membership or size checks.
4. Rerun all 46 scenarios and the two lifecycle checks, C0, ordinary FTP and
   test-enabled Release repetitions. The Move commitment map, IMAP buffering,
   real FTP/SSH session limits and wider route matrix remain separate open work.

The [original Copy baseline](../../Continuation/2026-09-05_c0_curl_copy_traversal_red/README.md),
[iterative candidate](../2026-09-05_c0_curl_copy_traversal_candidate_debug/README.md),
[broader candidate](../2026-09-05_c0_curl_copy_traversal_c0_debug/README.md),
[ordinary candidate regression](../2026-09-05_c0_curl_copy_traversal_ftp_regression_debug/README.md)
and [raw-code-23 diagnostic](../2026-09-05_c0_curl_copy_wide_diagnostic_debug/README.md)
remain immutable failure/continuation history, not competing execution checklists.

Raw perf: 20,981,209 bytes / 60,121 rows; SHA-256
`0d2c6626ec72ac5fc10c0c796ff4e9567150349f5593c1c8c652313c869972f0`.
`metric-selection.json` binds all 691 Curl/C0 rows in original order. Unrelated
high-frequency UI rows remain in the owned raw run; both full traces and all
results/reasons are retained. Source/localization: 197/197, 15.01 s. Spec inventory:
zero blocking findings. All 691 selected rows, original digest and order were
independently recomputed. Disk audit: 73 issues, none deleted. C0 and Done stay open.
