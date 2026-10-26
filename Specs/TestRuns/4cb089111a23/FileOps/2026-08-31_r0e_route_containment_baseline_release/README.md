# R0e route-containment baseline

- Machine profile: `4cb089111a23`
- Source commit: `1b0474f8` (R0e RED contracts, before production route classification)
- Build: test-enabled x64 Release `RedSalamander` target; receipt `7b4eedd5cdf14df1bb75a3b2bcfd53201f41cf44166b146fc2c1c3bcfd658288`; zero MSBuild warning/error diagnostics
- Command: `RedSalamander.exe --fileops-selftest --selftest-case=FileOps_ProviderCapabilityMatrix --perf`
- Runs: `r0e-baseline-release-final-20260831-01` through `-05`
- Scenario: five sequential independent processes, each performing 4,096 capability parses and 4,096 same-route admission decisions
- `FileOps.SelfTest.R0e.RouteContainmentUs`: `1,067,199`, `1,053,325`, `1,065,507`, `1,063,903`, `1,066,916` us
- Nearest-rank p95: `1,067,199` us
- Resource/telemetry shape: every sample reports 4,096 parses, 4,096 admission decisions, zero rejected-uncontained routes, and zero provider calls

Every baseline process is intentionally RED (`passed=2, failed=1, skipped=0`): Local's pre-change capability document has no `cancellation.routeClass`, so the bounded-route contract fails. The protected deterministic loop and all R0e metrics complete before that exact assertion; no synthetic baseline is substituted and the never-returning fixture is never invoked.
