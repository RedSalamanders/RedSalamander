# C0 native Move preflight — final Debug qualification

Command:
`Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_ -FailFast -RunId c0-curl-move-preflight-final-debug-20260905`.

**15/15 C0 cases pass**, zero skips. Native Move preflight passes 97/97 across all
32 scenarios; existing Delete guards (29), lost replies (33), partial-tree truth
(37), late listing (49), Delete bounds (127), and directory-size truth (85) pass.
This is focused evidence, not a Fresh Full repository verdict.

Build: 2:59, zero warnings/errors; suite 42,125 ms.
Receipt `f6037e70542e85072a3d0c51cd776a1cbbe3ca93fca97ab0ac72037c52fdfc99`;
code snapshot `7b36e420365f6c376e257c41a5e2a8b0bf05a41d1f2b8f0249f1400e31d8a4e2`.
Later closeout edits only consolidate specs/history and evidence.

Every malformed/out-of-root/overlong/depth-limit rejection precedes payload read
and destination/source mutation, with exact retained contents and one no-commit
receipt. Shallow/depth-80 controls succeed. All 4,097-file unknown-tail cases reach
the final SIZE probe and reject without mutation, not at an aggregate queue cap.
All four formerly unclean depth-129 routes now return buffer-overflow cleanly.
This does not establish the exact earlier process-termination stack.

Blocked-LIST cancel samples: **64,374 / 55,338 / 54,440 / 57,473 us**, versus about
1 second in the diagnostic Debug baseline. The fixture's emergency unblock is not
used. Deep controls take 997,227–1,143,163 us and wide-tail controls
1,222,672–1,291,463 us. Both are slower than the isolated diagnostic baseline
(801,072–880,414 / 944,498–993,523 us); process context differs and these are single
samples per route, not a throughput improvement or statistically qualified trend.

All 32 traversal/retained pairs stay within 129 frames, 81,024 path bytes and
6,613,056 metadata-reservation bytes. Those counters exclude the complete
commitment map, transient strings, caller inputs and opaque Curl/TLS/socket state.
Native Copy recursion/buffering, success at the admitted depth-128 boundary, map
storage and real-session quotas remain open.

Four non-perf runtime files are exact SHA-256-verified copies. The original
3,430,814-byte / 9,695-row perf stream remains in the owned run; the 446 exact
Curl/C0 rows here preserve raw order. `metric-selection.json` binds its digest and
original line numbers; it deliberately does not bind checkout-dependent newline
hashes of the selected file. The runner's 63 retained-run disk-audit findings
remain recorded; no old run was deleted to make that audit green.

[Earlier C0 history](earlier-c0-history.md) retains superseded plan execution detail
as non-normative history. Resume from the current WIP plan, not those old steps.

Closeout checks: source/localization Pester 197/197 (17.88 s); spec inventory zero
blockers; all 28 newly curated files pass explicit archive validation; complete
archive inventory passes 2,325 files. Both selection manifests were independently
recomputed (446/1,338 exact ordered rows), and all 40 local plan/history links
resolve. The production source SHA-256 still equals the tested snapshot digest
`5a2780e59a7338ba38804a3b7c30acee5f84d6c220c2e198f9ec54d35787ebe1`.
Patch/index whitespace checks pass. These checks qualify this saved repair, not
the still-open C0 route matrix or the whole WIP plan.
