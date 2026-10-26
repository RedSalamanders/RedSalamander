# Navigation failure diagnosis: e9099fb7-navigation-context

Local x64 Debug at `e9099fb751c3149f70b8e6104cdd980253936cea`; no hosted CI.
Each exact comma-separated case sequence in `summaries.json` runs ten times in one
Commands process. Original native results retain every case, repeat and failure reason.

| Sequence | Passed | Failed | Skipped | Run |
|---|---:|---:|---:|---|
| refresh-context | 130 | 0 | 0 | 20260913T101745Z-100884-241f347e1b9a479e9eeb792939784cd2 |
| drive-menu-context | 130 | 0 | 0 | 20260913T102057Z-100884-0526d5ce745d4390bcadba0a7713cfce |

These are focused diagnostic runs, not Fresh Full qualification. A passing reduced
sequence does not explain the earlier Full failure or prove that broader process state
is clean. Expanded context and short sequences must not be presented as matched performance
workloads. `case-comparison.json` extracts counts/failures without modifying original receipts.

`retained-files.json` hashes 8 original receipts. Complete native archives stay
under `Z:/RedSalamander.Perf/evidence/I19-20260909/e9099fb7-navigation-context`. The
[I19 checklist](../../../../Plans/Done/DxUi_SharedLibraryAdoptionAndReleasePlan_2026-09-09.md)
owns the remaining x64 qualification; ARM64/ASan are explicitly user-deferred.
