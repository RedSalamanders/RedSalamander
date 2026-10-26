# R1c cleanup-truth candidate

- Machine profile: `4cb089111a23`
- Source commit: `bfcc9cbe237e6d0251983dac49e0547505940329` (R1c implementation)
- Build: test-enabled x64 Release; full-solution receipt `77ad8b8d05fde6c971bdd4259b7e271d96e10d8048d01bd65e14c959bdce3ef7`; diagnostics `0 warning(s), 0 error(s)`
- Command: `Run-AllTests.ps1 -Suite FileOps -SkipBuild -Configuration Release -CaseFilter FileOps_ProviderCapabilityMatrix -TestRoot D:\RedSalamander.Perf -AllowAlternateVolumeTestRoot`
- Scenario: five sequential independent processes; each reduces 4,096 provider-free committed-primary results carrying retained cleanup debt
- `FileOps.CleanupDebt.ReduceUs`: `302`, `355`, `306`, `593`, and `380` us
- Nearest-rank p95: `593` us
- Admission budget: `max(1.50 * 569, 569 + 50,000) = 50,569` us
- Result: PASS (`593 <= 50,569`); every sample is below the `2,000,000` us absolute cap
- Resource shape: every process reports 4,096 Completed results, 4,096 independent Retained-debt facts, zero Indeterminate results, and 256 retained-state bytes per result

All five focused FileOps runs passed `3/0/0` with no skips. The exact implementation-tree Debug provider contracts also passed: Local `155/0`, Microsoft Drive `224/0`, S3 `195/0`, and S3 directory-transfer `157/0`; `PluginContractTests.exe` and `LocalizationTests.exe` both returned 0.
