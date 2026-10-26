# R7-A15 provider-name RED contract

- Machine profile: `4cb089111a23`
- RED test commit: `e610f3e4` (the run was captured immediately before committing the identical source guard)
- Production baseline commit: `ee6fdd875e94192d60e36fc7f2f1d04ed4da4ad3`
- Build: test-enabled x64 Release; full-solution receipt `421dc06995fd35877df5b2a9014767305ead8f2892ad8599c1fa1dff9e48031e`; diagnostics `0 warning(s), 0 error(s)`
- Run ID: `20260901T022207Z-81856-0b9f7d0412c84386ae8d7a7530df54f5`
- Command: `Run-AllTests.ps1 -Suite Commands -CaseFilter file_system_provider_name_policy_consumption_source_guard -Configuration Release -TestRoot D:\RedSalamander.Perf -AllowAlternateVolumeTestRoot -SkipBuild`
- Result: expected RED, `0` passed / `1` failed
- First blocking assertion: `R7-A15 requires one canonical composite child-name validation/join/collision result.`

This receipt proves the source guard distinguishes the activation baseline from the required R7-A15 production contract. Raw copies are retained under `D:\RedSalamander.Perf\evidence\r7-a15-provider-name-red`.
