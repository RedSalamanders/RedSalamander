HISTORICAL CHECKPOINT — superseded by the current I14 plan. Preserved for run links,
not an executable instruction or current status.

### C0 current evidence and next work — 2026-09-05

The receipt, replay-safety, partial-tree, native Delete, directory-size and Move
preflight-discovery repairs have focused qualification. C0 still remains **VERIFY**:
these passes are not the complete provider/route/walker matrix.

| Latest evidence | Result |
|---|---|
| [Debug C0](../2026-09-05_c0_curl_move_preflight_debug/README.md) | 15/15; all 97 Move preflight checks pass. |
| [Test-enabled Release C0](../2026-09-05_c0_curl_move_preflight_release_repeat3/README.md) | 45/45 across three repetitions; 97/97 preflight checks each time. |
| [Ordinary FTP workflow](../2026-09-05_c0_curl_move_preflight_ftp_regression_debug/README.md) | 3/3, same source snapshot as both final C0 profiles. |
| [First preflight candidate](../2026-09-05_c0_curl_move_preflight_first_debug/README.md) | 97/97; four explicit-special-member warnings subsequently fixed. |
| [Initial interrupted baseline](../../Continuation/2026-09-05_c0_curl_move_preflight_unclean/README.md), [diagnostic baseline](../../Continuation/2026-09-05_c0_curl_move_preflight_diagnostic/README.md) | Both retained. Out-of-root listing rows caused mutations in all four routes; depth-129 termination has no proven stack diagnosis. |

All three final builds are warning-free and bind code snapshot
`7b36e420365f6c376e257c41a5e2a8b0bf05a41d1f2b8f0249f1400e31d8a4e2`.
Preflight now uses the existing strict cursor and iterative stack, rejecting unsafe
rows/depth before any mutation while retaining per-file source-size commitments.
Blocked LIST cancellation is 54–64 ms Debug and 52–64 ms Release in these focused
loopback fixtures. Reservations peak at 129 frames, 81,024 path bytes and less than
6.62 MB of frame/cursor metadata; **the whole-tree commitment map is not bounded by
those metrics**. Deep/wide timings are retained without claiming throughput gains.

[Earlier C0 execution history](../2026-09-05_c0_curl_move_preflight_debug/earlier-c0-history.md)
preserves prior commands, receipts, passing runs and failures outside this execution
checklist. Its superseded next-step instructions are historical, not active work.
The enduring behavior remains in the owning VFS, File Operations, provider,
shared-helper and performance specs.

Native Copy/merge now shares one iterative strict-cursor walk with bounded serial
or scheduler-owned batches (`cf5b9b8e`). Size-probe metadata no longer goes to the
default CRT output (`61815385`). The fake endpoint now reuses consumed passive
listeners and rejects stale queued data connections; production code is unchanged
from `61815385` in the [latest C0 run](../2026-09-05_c0_curl_passive_lifecycle_debug/README.md).
That archive owns the full diagnosis, earlier-failure links and next test design.

Latest Debug: **15/16 harness cases pass** (13 named C0 cases plus Setup/Cleanup);
**Copy is 139/141 assertions**. All 46
native scenarios finish: 44 pass and both wide copies time out at the unchanged
120-second deadline during promotion (3,664/4,097 parallel; 3,005/4,097 serial).
Both endpoints are healthy; listener creations/reuses are 10/14,665 and 4/12,023.
Both new raw-client lifecycle checks, all deep/refusal/cancellation cases, and
[ordinary FTP 3/3](../2026-09-05_c0_curl_passive_lifecycle_ftp_regression_debug/README.md)
pass. Both builds are warning-free and bind snapshot
`685cd7bebb72f4be804e5e7287bb867fcd286d6de986503cfa1c252ca54b6805`.

**Next is cursor connection reuse, not Done.** Preserve the 4,097-file corpus,
deadline, full-listing/collision and source
size proofs. Frame/batch reservation peaks are not whole-operation memory: parent
lookups, bulk directory caches, the O(files) Move commitment map and opaque
transport/session state remain outside them. IMAP and real FTP/SSH session limits
also remain open; no test-enabled Release rerun qualifies this Copy revision.

The [new entry-lookup red baseline](../2026-09-05_c0_curl_entry_lookup_red/README.md)
now proves all eight single/bulk routes wrongly mutate for five unsafe parent
shapes (40 scenarios). `C0_CurlEntryLookupTruth` has 197 checks: 117 pass, 80 fail.
Literal/missing/cancel controls and both 16-query/4,097-row Unix/DOS probes pass.
The archive preserves fixture compile/oracle mistakes separately. Keep conservative
unknown-source receipts when lookup itself failed; zero requests alone cannot
prove originalStillPresent.

Strict bounded lookup is now implemented: the [focused candidate](../2026-09-05_c0_curl_entry_lookup_bounded/README.md)
passes 197/197, with 56–60 ms cancellation and 208/51,400 bytes of owned path/metadata
storage. Bulk parent caches and duplicate directory-size root lookup are removed.
However, the [same-source broad run](../2026-09-05_c0_curl_entry_lookup_connection_red/README.md)
is 15 pass / 1 fail / 1 not-reached: Copy 123/141. Repeated cursor multi destruction
churns connections; Tcpip 4227 confirms endpoint exhaustion. Repair pool ownership
before claiming this candidate qualified. The archive owns exact counts and diagnosis.

Continue in this order, committing each coherent code/test/spec/evidence step:

1. Resume the wide-copy diagnosis at step 3. For changed Curl paths, retain `C0_` and
   `R0fCurl_FakeFtpReadWriteCreateDelete` regressions. Use `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration <profile> -CaseFilter <case-or-prefix> -FailFast -RunId <new-id>`.
   Preserve/restore the process `RSBuildEnableTests` value around Release tests.
   Do not use `-SkipBuild` with a project-only receipt.
2. Freeze source/spec/index during governed runs. Archive immutable results/reasons,
   traces and digest-bound metrics; keep original failures. Run source/localization
   Pester, spec inventory, explicit/archive-inventory checks and `git diff --check`,
   then commit. First candidate guards passed 197/197 and spec inventory had zero
   blockers; repeat the applicable checks for the final revision.
3. Retain a cursor's multi connection cache with the existing bounded, exclusively
   borrowed easy-handle pool. Detach before return, reset borrowed callback state,
   preserve expiry/unload ownership and test cancellation/reborrow. Keep strict
   bounded lookup and full requested metadata. Rerun `C0_CurlEntryLookupTruth`
   (197 checks) and `C0_CurlCopyTraversalTruth` (141 checks, all 46 scenarios), retaining
   4,097 publications, depth boundaries and the two passive-lifecycle checks.
   Bound the separate Move commitment map without losing membership/size proof or
   imposing an aggregate 4,096-file cap. Retain preflight, size, Delete, receipt/replay
   and ordinary FTP regressions; finish with test-enabled Release repetitions.
4. Complete the provider/route/walker and host Rename matrix, then C1–C6. The four
   deferred extras remain HOLD. Exact Fresh Full, inherited I17, non-test Release
   SQLite staging and shared DLL Clean ownership still block Done.

Raw evidence stays beneath the initialized `C:/RedSalamander.Perf`. The untracked
repository-root `last_run/` is user-owned: never inspect, clean or stage it. Do not
delete retained runs to hide disk-audit findings. No push or destructive cleanup
is implied. Git metadata is outside this worktree at
`Z:/src/RedSalamander/.git/worktrees/RedSalamander-repair-staging`; staging and
commit use the normal approved escalation.
