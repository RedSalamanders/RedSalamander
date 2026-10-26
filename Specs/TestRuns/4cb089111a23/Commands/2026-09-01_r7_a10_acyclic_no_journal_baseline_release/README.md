# R7-A10 acyclic/no-journal baseline

- Machine profile: `4cb089111a23`.
- Source commit: `931ef726cda915ec10f128a06991624ff6f5be0a`
  (R7-A10 activation tree, before the production repair).
- Build: test-enabled x64 Release; full-solution receipt
  `d563efd4aec102a97ab248c5fdf18c5a2be300b16bcef80864811ba4a9b46074`;
  diagnostics `0 warning(s), 0 error(s)`.
- Command: `Run-AllTests.ps1 -Suite Commands -Configuration Release
  -SkipBuild -CaseFilter cmd_pane_batchRename_execution_engine_large_independent_perf`.
- Scenario: five independent processes execute the provider-free 1,024-row
  independent Batch Rename scheduler fixture.
- `batchrename.execute.us`: `197661`, `194087`, `193802`, `195024`, and
  `192742` us; nearest-rank p95 `197661` us.
- Candidate ceiling: `1.50 * 197661 = 296492` us, rounded to the nearest
  whole microsecond.
- Resource shape: every process completes 1,024/1,024 rows with zero failed
  rows. The legacy implementation has no accepted schedule/mutation/journal
  retirement metrics, so those invariants are evaluated only on the candidate.

All five baseline processes pass (`passed=1, failed=0, skipped=0`). Each
archived `perf_metrics.jsonl` retains the four claim-bearing execution rows in
their original order. The original runner JSONLs remain under the initialized
test root and are bound here as `(rows, bytes, SHA-256)`:

- sample01: `(320, 120523, 5C826690A13BF2B7A03CE7BC4CA121ACE61861CEE4BA2A425286FC54F1A1F963)`
- sample02: `(320, 120804, D5A272B90884470ABC7D89FC3EA086DB0B5A15F150A52CD91450A6AA90577590)`
- sample03: `(320, 120483, C1B9D7491B331C37A544B001C2DA29CC2D76F9E7B03299E64B249210E1F7B343)`
- sample04: `(320, 120483, E0E0593147B42432F291508F118D80E595CAA5C2C47A06456960C981F7A070F4)`
- sample05: `(320, 120495, 6AC54D6960075F1903EC0D7DEC21D98AF3188E62EBB0B552746183ACCE73EC27)`

The runner sandbox audit reports pre-existing content within the initialized
`D:\RedSalamander.Perf` root, including explicitly authorized sibling run
directories. It provides no evidence of a write outside that initialized root.
