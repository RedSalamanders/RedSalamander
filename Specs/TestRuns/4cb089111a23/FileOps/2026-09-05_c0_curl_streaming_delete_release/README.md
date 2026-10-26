# C0 bounded Curl Delete — test-enabled Release qualification

Same implementation/test/spec source snapshot as the passing Debug candidate:
`0522a18f77b9cfa47d372f211786011eb94bb52ae083455f806e1148d5351334`,
starting checkpoint `5132d866` plus recorded working changes.
Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Release -CaseFilter C0_ -FailFast -RunId c0-curl-streaming-delete-literal-release-20260905`,
with process-local `RSBuildEnableTests=true`, restored afterward.

Build: 3:55, zero warnings/errors. Receipt
`f1df3476948c4755bd0b09f5cf889e4c115d0bb3924fc8f61ed34b72d0109f19`.
Runtime: 13/13, zero skips, 17.6 seconds. Native assertions: guards 29/29,
lost replies 33/33, partial-tree/proof 37/37, late listing 49/49,
traversal/command operands 127/127. All 36 bounds scenarios succeed.

All 65 traversal/retained metric pairs satisfy the shared limits. Peaks:
129 ancestry frames including the depth-zero root, 128 queued leaves,
130,592 retained path bytes, 6,617,280 metadata reservation bytes. Serial queue
peaks are at most one; parallel peaks are at most 128. Text/metadata ceilings
are 16/8 MiB. Reservations exclude opaque TLS/socket state and process working set.
Configuration differences are not evidence of an optimization speedup.

Raw results/reasons, both traces, aggregate, and the entire ordered perf stream
are copied unchanged from the exact run-owned C:\RedSalamander.Perf artifacts.
The 1,813,875-byte stream fits the curated-file bound. The aggregate retains the
disk-audit limitations; older evidence was not removed.

This validates the focused native Delete slice, not live FTP/SSH, remaining
Copy/size walkers, C1–C6, or exact Fresh Full. It also does not close the earlier
non-test Release SQLite-staging or intermittent shared-output Clean findings.
