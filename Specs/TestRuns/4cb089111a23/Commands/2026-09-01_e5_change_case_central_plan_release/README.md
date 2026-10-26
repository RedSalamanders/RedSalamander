# E5 Change Case central-plan candidate

- Machine profile: `4cb089111a23`.
- Source: HEAD `bc160caeaebdb8c8133fd9c50bb2b66f242cfe27` plus the uncommitted E5 candidate.
- Build: test-enabled x64 Release; full-solution receipt
  `dd4c0d8597a8526d723ba44f9d42c0c8e1db3aaaf96096fa92dae6986d98b6e8`;
  diagnostics `0 warning(s), 0 error(s)`.
- Command: `Run-AllTests.ps1 -Suite Commands -Configuration Release
  -SkipBuild -CaseFilter cmd_pane_changeCase_central_rename_plan`.
- Scenario: five independent processes admit `RenameOrigin::ChangeCase`, build
  the immutable central rename plan, execute one exact Local case-only rename,
  publish one typed result, and verify the exact stored target casing.
- `changecase.admission.ui_thread_us` p95: `26216` us.
- `changecase.admission.wall_us` p95: `28696` us.
- `changecase.admission.worker_wall_us` p95: `2791` us.
- `changecase.execute.us` p95: `2378` us.
- Regression gate: PASS; every duration is below the `5000000` us scenario
  cap. Every process reports one accepted capability query, one exact bind, two
  provider-name queries, zero retained bindings, and one completed/published
  execution row with `hr=0`.

All five candidate processes pass (`passed=1, failed=0, skipped=0`). Each
archived `perf_metrics.jsonl` retains the nine claim-bearing Change Case
admission/execution rows in original order. The original runner JSONLs remain
under the initialized test root and are bound here as `(rows, bytes, SHA-256)`:

- sample01: `(529, 199262, AF0FBB4B9119A940DA224FEE419AE7E4DFFD3F0DC033EB61411796524B31E688)`
- sample02: `(529, 199407, 3EFFC9CC350D0B8D8C6C6503AEBA31C883CF0DB839A2AF908040C5B2AE1BF51A)`
- sample03: `(529, 199149, 6613DC0A82D0DC9897D5DD1CDCA7EC7B0391B11BF82B347F8B4C968BFD7C3B9F)`
- sample04: `(547, 205861, 8DD639619D75BCCCFB7ED11C720EC4CECE057BC147B947CE3B87D0F9F1E8488E)`
- sample05: `(529, 198927, B0042AF0AE49B41E128067C223091144AFFA4746DC392A29FBE267A508679CDB)`

The production candidate first built as a Debug full solution under receipt
`2e8d78b33dffd9454051c51907e00f667828ce532304e4e24fa1636b47c13318`.
After a test-only Windows exact-casing assertion correction, the exact-source
governed Debug product rebuild receipt is
`093b4fa4b32a4df8ce0a11a015b7db2d962a68ee03caaaff1ce0cddf8e1afa74`.
The runner sandbox audit reports pre-existing sibling runs only under the
explicitly initialized `D:\RedSalamander.Perf` root; it provides no evidence
of a write outside that root.
