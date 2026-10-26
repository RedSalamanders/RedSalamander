# R2 typed-route baseline

- Machine profile: `4cb089111a23`
- Source commit: `1c1ca225ed51d0baa99d026b3d6ea0d5f0339947` (R2 RED contracts)
- Build: test-enabled x64 Release; full-solution receipt `8eb9fa88e523972e079b3d88265f5bdc12171b7c7b6666f4dae6aec8b40f295d`; diagnostics `0 warning(s), 0 error(s)`
- Command: `Run-AllTests.ps1 -Suite FileOps -SkipBuild -Configuration Release -CaseFilter FileOps_ProviderCapabilityMatrix -TestRoot D:\RedSalamander.Perf -AllowAlternateVolumeTestRoot`
- Scenario: five sequential independent processes; each performs 4,096 offline Dummy capability queries and parses the legacy JSON document
- `FileOps.RouteFacts.QueryUs`: `4,146`, `4,357`, `4,163`, `4,612`, and `4,782` us
- Nearest-rank p95: `4,782` us
- Resource shape: every process reports 4,096 queries, 4,096 JSON parses, zero typed queries, zero arena fallbacks/rejections, and 688,128 cumulative result bytes (168 bytes per snapshot)

Every baseline process is intentionally RED (`passed=2, failed=1, skipped=0`). The exact failure is that the shipped Dummy provider does not expose the new typed route-capability IID; the measurements are emitted before that assertion. The separately retained Commands RED case fails because the new interface/spec authority and production JSON-use ban are not yet present.
