# R1c cleanup-truth baseline

- Machine profile: `4cb089111a23`
- Source commit: `c20eed1c3b0f223dda3ba440114c5ab7071da02d` (R1c RED contract)
- Build: test-enabled x64 Release; full-solution receipt `29cc49e75f99578208d1a8037f4f53913a9d5e4c41e66d5fe563e5d373b5550d`; diagnostics `0 warning(s), 0 error(s)`
- Command: `Run-AllTests.ps1 -Suite FileOps -SkipBuild -Configuration Release -CaseFilter FileOps_ProviderCapabilityMatrix -TestRoot D:\RedSalamander.Perf -AllowAlternateVolumeTestRoot`
- Scenario: five sequential independent processes; each reduces 4,096 provider-free committed-primary results carrying retained cleanup debt
- `FileOps.CleanupDebt.ReduceUs`: `569`, `319`, `351`, `435`, and `457` us
- Nearest-rank p95: `569` us
- Resource shape: every process reports 4,096 results, zero Completed, 4,096 Retained, 4,096 Indeterminate, and 256 retained-state bytes per result

Every baseline process is intentionally RED (`passed=2, failed=1, skipped=0`). The legacy host reduction collapses committed primary success plus `Retained` cleanup debt into `Indeterminate`; the exact failure states that committed results must remain Completed while preserving the independent retained-debt axis.
