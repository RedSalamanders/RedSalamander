# Bounded cursor release: Debug serial wide copy remains red

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild
-TestRoot C:\RedSalamander.Perf -CaseFilter C0_
-RunId c0-curl-cursor-release-broad-debug-20260906`.

Result: 15 passed, 1 failed, 1 skipped/not reached; 244.6 seconds runner duration.
Copy: 141/142 checks across all 46 scenarios, 205,984 ms. Broad lookup is explicitly
not reached after failure; its separate exact run passes 284/284. Delete 127/127,
directory size 85/85 and Move preflight 97/97 pass. All depth-128 Copy/Move routes
pass with two authenticated sessions per endpoint. No connection-failure metrics
occur in this run. This is not broad C0 qualification.

| Wide Copy | Prior listing candidate | Cursor-release candidate |
|---|---|---|
| Parallel, concurrency 4 | 4,097/4,097; 71.420825 s | 4,097/4,097; 64.002003 s |
| Serial, concurrency 1 | 3,828/4,097; timeout at 119.987399 s | 4,041/4,097; timeout at 119.997810 s |

Baseline: `../2026-09-05_c0_curl_listing_broad_debug/README.md`, same machine.
Serial completion differs, so elapsed time is not an equal-work throughput
comparison. Neither run has overlapping builds/policy checks or OS socket probes.
The timeout, corpus, concurrency and assertions are unchanged. Both candidate
wide scenarios finish with zero pending passive-listener retirements (one historical
retirement event each); no fixture ownership guard is bypassed.

The remaining cost is concentrated in serial staging-name lookup (51,907,702 us,
4,042 visits) and promotion (54,154,806 us, 4,041 visits). Destination LIST generation
plus its state-lock wait accounts for 39,371,795 us across 8,084 LISTs and
849,527,411 bytes. These nested timings must not be added to client wall time.
Parallel worker phase times overlap. Next work should reduce repeated strict
listing/row materialization cost while retaining full malformed/duplicate/missing/
timestamp/final-reply validation; do not raise the deadline, truncate listings,
cache stale mutation authority or add automatic replay.

The receipt matches the sibling Debug lookup build and frozen Release source.
Exact copies and ordered selected metrics carry raw digests and original row
positions. `prior-checkpoint.md` is an exact snapshot of the accumulated WIP plan
before its diagnostic history was condensed; `plan-snapshot.json` binds its raw
bytes to the tested source snapshot. Historic relative links in that snapshot
retain their original WIP base and are not newly relocated normative references.

The connection-lifetime fix is retained because nested transport reuse and deep
operations are proved; it is not claimed to fix this separate wide timeout.
Whole-operation Move memory, IMAP policy/buffering, real endpoint quotas, the route
matrix and C1-C6 still block moving the core plan to Done.
