# Broader C0 after iterative Copy — 15 passed, 1 failed

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_ -FailFast -RunId c0-curl-copy-traversal-c0-debug-20260905`.

Build 02:29, zero warnings/errors. Receipt:
`24c9863439e3d909897447dea4595cae8feb85c9b7c7c9f86e2dc83be954a041`.
Workspace `2d8e0ab053fb908e6580526cc73f650658015426fc14bf0c969ee69e44db2196`
matches the [first candidate](../2026-09-05_c0_curl_copy_traversal_candidate_debug/README.md).

All previous 15 cases pass; the new Copy gate fails (135/139 assertions, 44/46
scenarios pass). No skips. Suite duration 190.8 s; Copy case 147.3 s. Native guards
retain 29/29, lost mutation reply 33/33, partial-tree 37/37, Delete late-listing
49/49, Delete bounds 127/127, directory-size 85/85 and Move preflight 97/97.
Depth-128 Copy/Move and depth-129 refusal pass again without process termination.

Remaining wide failures reproduce: serial ERROR_TIMEOUT after 120,011,282 us and
2,937 publications (source idle-control ERROR_SEM_TIMEOUT); parallel source
proof/download E_FAIL after 233,705 us and 55 publications. Full-source absence of
deletes is not a completed Copy; no aggregate-count cap or shortened corpus passed.
Blocked first/late LIST cancellation is 50,942–66,369 us. Peak admitted Copy
reservations are again 129 frames, 128 queued files, 226,472 path bytes and
6,670,544 metadata bytes; separate lookup/map/transport allocations are excluded.

During the serial wide scenario a read-only elevated TCP query observed 14,213
loopback TIME_WAIT connections, 12,624 distinct local ports and 8,038 distinct
remote ports. The configured dynamic TCP range is 49,152–65,535 (16,384 ports).
This suggests investigating fixture transport pressure, **not a proven diagnosis**
of the generic E_FAIL. An earlier non-elevated query was access-denied; its empty
result is not zero-usage evidence. No TCP settings or unrelated processes changed.

Next diagnosis:

1. Add content-free CURLcode/response-code diagnostics and distinguish source SIZE
   proof from payload download. Isolate the parallel wide case from the preceding
   120-second serial case to establish whether order/resource pressure matters.
2. Inspect the fake server's per-transfer passive-listener churn and 30-second
   control-idle policy. If reusing a listener, preserve transfer ownership and
   handle stale queued connections after denied commands. Do not change global
   TCP settings or hide fixture errors to obtain green.
3. Bound/qualify parent admission and publication lookups: `GetEntryInfo` still
   buffers whole directories; `GenerateRemoteSiblingPath` and overwrite preparation
   repeatedly list the growing destination. Do not discard absence, collision or
   size proof. Keep all 4,097 real files; justify any workload-specific time budget
   explicitly against measured work instead of silently increasing it.
4. Complete bounded Move commitment storage and remaining route matrix; rerun full
   new gate, C0, ordinary FTP and test-enabled Release repetitions before claiming
   this component qualified. Exact Fresh Full still governs overall Done.

Raw perf is 12,989,508 bytes / 37,142 rows. `metric-selection.json` binds its SHA-256
and all 674 Curl/C0 rows in original order; unrelated high-frequency UI polling
rows remain in the immutable raw run beneath `C:/RedSalamander.Perf`. Full original
results/reasons and both traces are retained. Disk audit: 68 issues, not cleaned.

Checkpoint checks: source/localization Pester 197/197; final updated-plan spec
inventory zero blockers; all 19 files in the three candidate archives pass the
explicit archive gate. All 674 selected metric rows and their order were
independently recomputed; 13 local plan links resolve. Production source SHA-256
matches the final FTP tested snapshot. Working diff whitespace check is clean.
