# R1b Change Case truth baseline

- Machine profile: `4cb089111a23`.
- Source commit: `529eca0f8c68f5071531c0bf20ad2ae8fc512491`
  (R1b RED tree, before the production repair).
- Build: test-enabled x64 Release; full-solution receipt
  `672af8c60958e6bb3c9afed18ff2f1db98ddb7ef0d463882679d43336d0fc2ba`;
  diagnostics `0 warning(s), 0 error(s)`.
- Command: `Run-AllTests.ps1 -Suite Commands -Configuration Release
  -SkipBuild -FailFast -CaseFilter cmd_pane_changeCase_task_payload_truth`.
- Scenario: five independent processes dispatch and reduce 4,096 Change Case
  task updates through the legacy receipt fixture, with no provider call.
- `Commands.SelfTest.R1b.ChangeCaseTaskDispatchUs`: `160`, `156`, `154`,
  `172`, and `155` us; nearest-rank p95 `172` us.
- Resource shape: every process reports 4,096 payloads, 4,096 task creates,
  zero task updates, zero provider calls, and zero outstanding receipts.

Every baseline process is intentionally RED (`passed=0, failed=1, skipped=0`)
because it created 4,096 simulated task identities instead of one. Each
archived `perf_metrics.jsonl` retains the six claim-bearing R1b rows in their
original order. The original runner JSONLs remain under the initialized test
root and are bound here as `(rows, bytes, SHA-256)`:

- sample01: `(326, 122832, 43DFE297EFE81EFF6CA68104E15F81A8AF4D27470039E64DFE60727E4E2929F8)`
- sample02: `(326, 122830, F35D14EA47D015ABB46258DA1908D26973179B48F0E948EFB27A857679694748)`
- sample03: `(326, 122845, 070245FB1FC85ABD3CA83EA4B119CBF2E02F1E0FE72A51B5565D3416042449E2)`
- sample04: `(326, 122812, C9AF7288AB4D7E99DAAC9DDEF9BBE25B8EF9216807203D93F62DCB1DAA315A42)`
- sample05: `(326, 122850, 701C35637C875899FB8F7DFCF70751C8A40DF5CE1A5DD6BF8A15E0067CFC13DA)`

The runner's sandbox audit reports earlier explicitly authorized sibling runs
under `D:\RedSalamander.Perf`; it found no write outside that initialized root.
