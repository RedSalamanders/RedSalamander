# Bounded cursor release: broad Release C0

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild
-TestRoot C:\RedSalamander.Perf -CaseFilter C0_
-RunId c0-curl-cursor-release-broad-release-20260906`.

Result: 17 passed, zero failed/skipped, 96.4 seconds runner duration (95,532 ms
native suite). Copy 142/142 across all 46 scenarios; lookup 284/284 is reached and
passes. Delete 127/127, directory size 85/85 and Move preflight 97/97 pass.
All four depth-128 Move variants complete with exact source/destination contents
and truthful completion. Each deep Copy/Move route records two authenticated
sessions at each endpoint. There are no Windows 10048 events; the six 10061
connection-failure rows and two socket trace rows belong to the intentional
refused-endpoint lookup controls after Copy, not operation failures.

Both wide copies publish exactly 4,097/4,097 files, with zero pending passive
listener retirements: parallel 6,219,373 us; serial 18,936,925 us. The same-machine
exclusive-listener baseline in
`../2026-09-06_c0_curl_exclusive_broad_release/` was 134/142 with four failed
deep Moves, and wide times 7,426,246 / 23,920,633 us. This is a single-run
directional timing comparison, not a throughput guarantee. Workloads, assertion
counts, concurrency, timeout scale and operation deadlines are unchanged.
The candidate's native Copy case is 43,969 ms versus 50,250 ms in that baseline;
whole-suite time is not comparable because the old failure prevented lookup.

Source/build match the sibling lookup Release archive, which retains the complete
tested patch and zero-warning full-solution build. Exact copies and all selected
Curl/C0 metrics preserve raw hashes and original row order through the two
provenance manifests. Native case statuses, including intentional faults, are
retained rather than inferred from the aggregate verdict.

This establishes this broad Release regression, not the entire C0 route matrix,
Debug qualification, live mutations, provider quotas or remaining C1-C6 work.
