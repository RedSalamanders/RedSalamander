# E5 Change Case discovery candidate

- Machine profile: `4cb089111a23`.
- Source: HEAD `bc160caeaebdb8c8133fd9c50bb2b66f242cfe27` plus the uncommitted E5 candidate.
- Build: test-enabled x64 Release; full-solution receipt
  `dd4c0d8597a8526d723ba44f9d42c0c8e1db3aaaf96096fa92dae6986d98b6e8`;
  diagnostics `0 warning(s), 0 error(s)`.
- Command: `Run-AllTests.ps1 -Suite Commands -Configuration Release
  -SkipBuild -CaseFilter cmd_pane_changeCase`.
- Scenario: five independent processes discover and type 1,024 Local rename rows
  without mutating through the legacy Change Case path.
- `Commands.SelfTest.E5.ChangeCaseDiscoveryUs`: `3118916`, `3206236`,
  `3060033`, `2813246`, and `3116779` us; nearest-rank p95 `3206236` us.
- Regression gate: PASS (`3206236 <= 5000000`). Every process reports
  `value0=1024`, `value1=1024`, and `hr=0`.

All five candidate processes pass (`passed=1, failed=0, skipped=0`). Each
archived `perf_metrics.jsonl` retains the one claim-bearing discovery row. The
original runner JSONLs remain under the initialized test root and are bound
here as `(rows, bytes, SHA-256)`:

- sample01: `(331, 125664, F2403547B36C6A71956EA38C3CCE2C0E5A97C550C8338BCA84DBE97B70119727)`
- sample02: `(331, 125678, 432EA210B0530DAB94C1E369D2F559D6AEDF911F10FEE1B395B8F888952AC59E)`
- sample03: `(349, 132268, 964420210965E5AB59174A418F2D3981FCA784EF7DCF247C8530A95ED9693816)`
- sample04: `(331, 125918, E35A692732ED969C624D0C01B0B44827251B844F9C028B5477DD018B792CD94E)`
- sample05: `(331, 125977, 2A336C51CDBF8581220B9E83F853A594EFD113C4178F33EAAF6DC3EF7DBE2BAF)`

The production candidate first built as a Debug full solution under receipt
`2e8d78b33dffd9454051c51907e00f667828ce532304e4e24fa1636b47c13318`.
After a test-only Windows exact-casing assertion correction, the exact-source
governed Debug product rebuild receipt is
`093b4fa4b32a4df8ce0a11a015b7db2d962a68ee03caaaff1ce0cddf8e1afa74`.
The runner sandbox audit reports pre-existing sibling runs only under the
explicitly initialized `D:\RedSalamander.Perf` root; it provides no evidence
of a write outside that root.
