# Focused adoption regressions: a8350f4b-x64-Debug-focused

Local x64 Debug at `a8350f4bf3e8d18cd666c720e2628167b4eb3af3`; no hosted CI.
Each declared case sequence in `summaries.json` runs ten times in its native self-test
suite. File Operations totals include setup/cleanup cases. Original native results
retain every case, repeat and failure reason.

| Sequence | Passed | Failed | Skipped | Run |
|---|---:|---:|---:|---|
| refresh-context | 130 | 0 | 0 | 20260913T103414Z-66012-3af9a5e14bbb455aa1b9e45f46264cf1 |

These are focused diagnostic runs, not Fresh Full qualification. A passing reduced
sequence does not explain the earlier Full failure or prove that broader process state
is clean. Expanded context and short sequences must not be presented as matched performance
workloads. `case-comparison.json` extracts counts/failures without modifying original receipts.

`retained-files.json` hashes 5 original receipts. Complete native archives stay
under `Z:/RedSalamander.Perf/evidence/I19-20260909/a8350f4b-x64-Debug-focused`. The
[I19 checklist](../../../../Plans/Done/DxUi_SharedLibraryAdoptionAndReleasePlan_2026-09-09.md)
owns the remaining x64 qualification; ARM64/ASan are explicitly user-deferred.

## Refresh fixture comparison

The same thirteen-case context passes 130/130 before (`e9099fb7`) and after
(`a8350f4b`) the quiet-baseline correction. The ten target Refresh cases pass in
each run. `refresh-baseline-review.json` records their whole-case durations and
proves that the Refresh dispatch and every following assertion are byte-identical
in Git. The additional quiet-baseline observation belongs to fixture setup; these
durations do not measure user-facing Refresh latency or change an application budget.
All 14 actual consumer module identities match after the corrected context run.
The earlier Full failure remains unexplained by these passing context runs.
