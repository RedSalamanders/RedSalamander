# R1b Change Case truth candidate

- Machine profile: `4cb089111a23`.
- Source commit: `b9df3a3a21d90d2c109cb5fc1dcf3b5d98f7e116`.
- Build: test-enabled x64 Release; full-solution receipt
  `38841fe2388d7d8522fa0cb0c7d5a7b3d1c0f211623b4780c626f85367973024`;
  diagnostics `0 warning(s), 0 error(s)`.
- Command: `Run-AllTests.ps1 -Suite Commands -Configuration Release
  -SkipBuild -FailFast -CaseFilter cmd_pane_changeCase_task_payload_truth`.
- Scenario: five independent processes dispatch and reduce 4,096 tokenized
  Change Case task updates through one shared creation receipt, with no
  provider call.
- `Commands.SelfTest.R1b.ChangeCaseTaskDispatchUs`: `172`, `172`, `172`,
  `172`, and `173` us; nearest-rank p95 `173` us.
- Admission ceiling: `max(1.50 * 172, 172 + 50,000) = 50,172` us. Result:
  PASS (`173 <= 50,172`); every sample is below the `2,000,000` us cap.
- Resource shape: every process reports 4,096 payloads, one task create, 4,095
  task updates, zero provider calls, and zero outstanding receipts after drain.

All five candidate processes pass (`passed=1, failed=0, skipped=0`). Each
archived `perf_metrics.jsonl` retains the six claim-bearing R1b rows in their
original order. The original runner JSONLs remain under the initialized test
root and are bound here as `(rows, bytes, SHA-256)`:

- sample01: `(389, 146091, C6D20E6317E35365A59EFC2E90197BD914FDB36CD524D45E49930F3A42423EB6)`
- sample02: `(389, 146081, 91C50E4BA8B1FACC5D360335F1D4FD3ACA2A3FCB020291705154DC7AAF466CB1)`
- sample03: `(389, 146055, 4F01D3995FB755AB95FA4E5A4F034FBB30032DB4A9C7442C289E89C0AEDF107B)`
- sample04: `(389, 146041, 4CC60E1752F7C36D8CD9479A977A48EB1F74837E671882CA22C1F51C58CB6223)`
- sample05: `(389, 146016, 95E6CE2C3A8AD6AEFFFE6CB054E15657DACAD473CC43E366AB3731439C9B06BE)`

The exact-commit Debug build receipt is
`8ea7cb1a750830121765430a7d97795083ee3a13d82006cb80b6e4f863ab56ef`.
The three new focused Debug cases and both retained Change Case Release cases
pass, the Commands source-contract suite passes `186/186`, and the corrected
Release `LocalizationTests.exe` invocation passes. The runner's sandbox audit
reports earlier explicitly authorized sibling runs under
`D:\RedSalamander.Perf`; it found no write outside that initialized root.
