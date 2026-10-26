# R0e route-containment candidate

- Machine profile: `4cb089111a23`
- Source commit: `bdacb1ea`
- Build: test-enabled x64 Release full solution; receipt `ceea32e11a7572cdad338e358941a1b16fbcdeb243cf77e98860ad461ba3deec`; diagnostics `0 warning(s), 0 error(s)`
- Debug full-solution receipt: `c667ef40a249434fb765d6cfc47229d594d8a10894461f9aec3d604e17a08acb`; diagnostics `0 warning(s), 0 error(s)`
- Command: `RedSalamander.exe --fileops-selftest --selftest-case=FileOps_ProviderCapabilityMatrix --perf`
- Runs: `r0e-candidate-release-final-20260831-01` through `-05`
- Scenario: five sequential independent processes, each performing 4,096 capability parses and 4,096 same-route admission decisions
- `FileOps.SelfTest.R0e.RouteContainmentUs`: `1,090,078`, `1,092,912`, `1,076,248`, `1,055,543`, `1,075,131` us
- Nearest-rank p95: `1,092,912` us
- Admission budget: `max(1.50 * 1,067,199, 1,067,199 + 50,000) = 1,600,799` us
- Result: PASS (`1,092,912 <= 1,600,799`); every sample is below the `2,000,000` us absolute cap
- Resource/telemetry shape: every sample reports 4,096 parses, 4,096 admission decisions, one rejected-uncontained witness, and zero provider calls

All five candidate processes passed `3/0/0` with no skips. The uncontained never-returning fake is rejected with `ERROR_NOT_SUPPORTED` before task publication, worker creation, or a provider operation call. Local fixed-volume routes remain bounded; synthetic UNC is always classified as `local-win32-smb`/uncontained without opening the share; MTP remains provider-watchdog qualified. Exact-commit qualification also passed the Debug FileOps matrix `3/0/0`, `mtp_capabilities_are_instance_honest` `1/0/0`, and Release `PluginContractTests.exe` including Microsoft Drive `223/0`, Curl `226/0`, S3 multipart `50/0`, S3 directory probes `157/0`, and unload quiet points. This is in-process capability/admission containment evidence, not live-network, SMB-throughput, or generic reader-cancellation evidence.
