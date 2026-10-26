# C0 bounded Curl Delete — Debug qualification

Starting checkpoint `5132d866` plus the bounded cursor/iterative Delete, literal
command-operand fix, fixture corrections, tests and owning spec updates.
Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_ -FailFast -RunId c0-curl-streaming-delete-literal-debug-20260905`.

Build: 2:45, zero warnings/errors. Receipt
`ff32fe87f570ef1f86dc019a19cd4714c3ab701d98bc3a9e409905265989246a`;
source snapshot `0522a18f77b9cfa47d372f211786011eb94bb52ae083455f806e1148d5351334`.
Runtime: 13/13, zero skips, 22.7 seconds. Native assertions: guards 29/29,
lost replies 33/33, partial-tree/proof 37/37, late listing 49/49,
traversal/operand checks 127/127.

The 36 bounds scenarios cover 4,097 files, 80 levels, explicit depth overflow,
literal arrow names, no-follow link entries, malformed/outside/overlong records,
and nested spaces/quotes/percent/Unicode names. Eighteen additional pure operand
checks cover FTP/SFTP/SCP literal base paths, quoting, percent identity, and
CR/LF/NUL/invalid-Unicode refusal. These are not live SSH protocol tests.

All 65 traversal/retained metric pairs were checked: maximum 129 frames (root at
depth zero), 128 queued leaves, 130,592 retained path bytes, and 6,619,344 metadata
reservation bytes. Every concurrency-1 tree has at most one queued leaf; the
parallel target is at most 128. Shared path/metadata ceilings are 16/8 MiB.
These are owned-text/reservation counters, not process working set or opaque
TLS/socket memory. Wide success must not fail at the 4,096 aggregate-entry boundary.

The late-listing baseline at `5132d866` passed 31/49 assertions; all six
concurrency-4 cases stalled before dispatching earlier work. This candidate passes
all twelve ordinary/late-failure/cancel scenarios. Earlier mutations remain
indeterminate after failure/cancel; untouched independent selected roots keep
their exact access-denied cause and no-commit proof.

Raw results/reasons and both traces are copied unchanged. The full perf stream
exceeds the 2-MiB curated-file limit and remains at the exact run-owned test-root
path. `metric-selection.json` binds its byte/row counts and SHA-256, deterministic
case-insensitive metric-name filter, and original line numbers. The compact
`perf_metrics.jsonl` preserves all 218 selected rows in original order; it is not
a whole-run throughput/percentile summary. Disk-audit limitations remain explicit
in the aggregate; no prior evidence was removed.

This qualifies the focused native Delete slice only. Copy/merge/size walkers,
live FTP/SSH session/resource limits, broader host journeys, C1–C6, and exact Fresh
Full remain separate gates. Do not move the plan to Done from this archive alone.
