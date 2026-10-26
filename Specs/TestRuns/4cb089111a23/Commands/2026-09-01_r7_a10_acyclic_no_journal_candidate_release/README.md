# R7-A10 acyclic/no-journal candidate

- Machine profile: `4cb089111a23`.
- Source commit: `8c303593aa30a30b892967d8fc642db492231ca1`.
- Build: test-enabled x64 Release; full-solution receipt
  `852d50aba0b168a721c6daa2f7bfd59275cab5a8dca04dc2535581c192df874d`;
  diagnostics `0 warning(s), 0 error(s)`.
- Command: `Run-AllTests.ps1 -Suite Commands -Configuration Release
  -SkipBuild -CaseFilter cmd_pane_batchRename_execution_engine_large_independent_perf`.
- Scenario: five independent processes build one immutable acyclic schedule
  and execute 1,024 provider-free final mutations.
- `batchrename.schedule.build.us`: `1257`, `1292`, `1212`, `1281`, and
  `1258` us; nearest-rank p95 `1292` us.
- `batchrename.execute.us`: `1801`, `1832`, `1723`, `1790`, and `1776` us;
  nearest-rank p95 `1832` us.
- Regression gate: PASS (`1832 <= 296492`); every schedule and execution
  duration is below the `5,000,000` us scenario cap.
- Resource shape: every process reports one layer, zero cycle members,
  `106496` retained index bytes, 1,024 rows/completions/mutations, zero failed
  rows, and zero journal writes.

All five candidate processes pass (`passed=1, failed=0, skipped=0`). Each
archived `perf_metrics.jsonl` retains the ten claim-bearing schedule/execution
rows in their original order. The original runner JSONLs remain under the
initialized test root and are bound here as `(rows, bytes, SHA-256)`:

- sample01: `(326, 122760, 658DB7B773AFE9291B3AEE5252644B7A902690A4D686C42FBB41A6D6C4670E19)`
- sample02: `(326, 122763, 43D5324E1A7261D9F9FC4AE56ED38BE75C2AB6AE6B6C2AB9D93C885C799CF393)`
- sample03: `(326, 122764, B1C0DF5A40F1ACC2CFD4D0CA08BDF1518EC3CEE5D13F1C9F1E939655C339B836)`
- sample04: `(326, 122744, 44C440F40C8AB8AE794D850BF894ACEECC2D53272D8F142604BB8544F2E7C260)`
- sample05: `(326, 122760, 4F4D5FD0FAFFAC323BB3BFE2DCD5E8B155D7466E604B61605A85D488066B6FFA)`

The exact-commit Debug full-solution receipt is
`67b1cf25cee85079c9581cb360fba1225558cd1da222d44395eee749553322a6`.
The cycle-preview, worker-rejection, direct-engine, no-journal, chain,
directory-chain, case-only, and spec-contract Release cases pass individually;
`LocalizationTests.exe` also passes. The runner sandbox audit reports
pre-existing content within `D:\RedSalamander.Perf`, including explicitly
authorized sibling run directories. It provides no evidence of a write outside
that initialized root.
