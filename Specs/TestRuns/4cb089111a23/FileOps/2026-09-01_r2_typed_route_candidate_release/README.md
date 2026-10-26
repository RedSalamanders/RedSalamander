# R2 typed-route candidate

- Machine profile: `4cb089111a23`
- Source commit: `75df3717` (R2 typed-route implementation)
- Build: test-enabled x64 Release; full-solution receipt `ac0e6f452203c56bb71f9bc235828a725524097aefd8f41d44022d29034680fd`; diagnostics `0 warning(s), 0 error(s)`
- Command: `Run-AllTests.ps1 -Suite FileOps -SkipBuild -Configuration Release -CaseFilter FileOps_ProviderCapabilityMatrix -TestRoot D:\RedSalamander.Perf -AllowAlternateVolumeTestRoot`
- Scenario: five sequential independent processes; each performs 4,096 provider-free typed route queries over the offline Dummy profile
- `FileOps.RouteFacts.QueryUs`: `1,450`, `1,432`, `1,437`, `1,440`, and `1,462` us
- Nearest-rank p95: `1,462` us
- Baseline p95: `4,782` us; acceptance ceiling: `54,782` us (`max(1.50 * baseline p95, baseline p95 + 50,000 us)`)
- Resource shape: every process reports 4,096 queries, zero JSON parses, 4,096 typed queries, zero arena fallbacks, zero rejections, and 917,504 cumulative result bytes (224 bytes per typed snapshot)

Every candidate process passed all three focused cases. Candidate p95 is 69.4% below the same-machine RED baseline, every individual sample is below 2,000,000 us, and the query counts and resource-shape counters are constant across all five processes.

The checked-in JSONL files retain the seven protected `FileOps.RouteFacts.*` rows from each process. The unfiltered raw captures, including unrelated startup and rendering telemetry, remain under `D:\RedSalamander.Perf\evidence\r2-typed-route-candidate-release\sample01` through `sample05` on the authorized test volume.
