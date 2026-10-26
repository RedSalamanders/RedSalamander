# Focused adoption regressions: 1232d181-history-context

Local x64 Debug at `1232d181e74ae7118d3fce7974196811b3cd4398`; no hosted CI.
Each declared case sequence in `summaries.json` runs ten times in its native self-test
suite. File Operations totals include setup/cleanup cases. Original native results
retain every case, repeat and failure reason.

| Sequence | Passed | Failed | Skipped | Run |
|---|---:|---:|---:|---|
| history-context | 130 | 0 | 0 | 20260913T131547Z-38000-92e9fb4e33614d4eac7c75686e48bbe2 |

These are focused diagnostic runs, not Fresh Full qualification. A passing reduced
sequence does not explain the earlier Full failure or prove that broader process state
is clean. Expanded context and short sequences must not be presented as matched performance
workloads. `case-comparison.json` extracts counts/failures without modifying original receipts.

`retained-files.json` hashes 12 original receipts. Complete native archives stay
under `Z:/RedSalamander.Perf/evidence/I19-20260909/1232d181-history-context`. The
[I19 checklist](../../../../Plans/Done/DxUi_SharedLibraryAdoptionAndReleasePlan_2026-09-09.md)
owns the remaining x64 qualification; ARM64/ASan are explicitly user-deferred.

The root build passes in 4m15s with zero warnings/errors. The ancestor review and
original source-check receipts describe the intermediate `e61d5e34` fixture;
`ancestor-syntax-review.json` proves that `1232d181` removes only its extra closing
parenthesis. Final module identities bind the actual tested candidate. The unchanged
post-click predicate, nine bounded deadlines and ancestor navigation/cleanup suffix
remain protected. Extra settling time belongs to fixture setup, not application latency.

The retained matched-sequence timings are in `ancestor-duration-comparison.json`.
Ancestor-case median elapsed time is 96.5 ms before and 423.5 ms after (+327 ms).
Both sequences pass 130/130. The added stable-geometry observations are fixture
setup; these whole-case timings also include initialization and scheduling. They
do not measure application response latency or establish the original failure's cause.
No production source under Common, Plugins or the application changes between the
two revisions; the only native change is the Commands fixture source.
