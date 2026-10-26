# R1a terminal and clipboard truth candidate

- Machine profile: `4cb089111a23`
- Source commit: `c2e055e9` (R1a implementation)
- Build: test-enabled x64 Release; full-solution receipt `269b700e5705d5105646a1994ebcc4a0395837379703b5188ca3abc5619da1b0`; diagnostics `0 warning(s), 0 error(s)`
- Command: `RedSalamander.exe --fileops-selftest --selftest-case=Phase10_ClipboardAdmissionAndRetainedActions --selftest-fail-fast --selftest-no-activate`
- Scenario: five sequential independent processes; each initializes/finalizes 4,096 selected-root result builders, attempts one duplicate terminal store, and evaluates 1,024 no-provider readiness/clipboard decisions
- `FileOps.SelfTest.R1a.TerminalFunnelUs`: `769`, `1,174`, `558`, `608`, and `661` us
- Nearest-rank p95: `1,174` us
- Admission budget: `max(1.50 * 619, 619 + 50,000) = 50,619` us
- Result: PASS (`1,174 <= 50,619`); every sample is below the `2,000,000` us absolute cap
- Resource shape: every sample reports 4,096/4,096 builders/immutable terminal stores, `1/1` duplicate rejection, 1,024 decisions with exactly 512 mutation permits, and zero provider calls; process-local fixture state is released at process exit

All five focused FileOps runs passed `3/0/0` with no skips. The exact full-build Debug case also passed `3/0/0`, the provider-capability/receipt matrix passed `3/0/0`, and `LocalizationTests.exe` returned 0. A broad Debug FileOps run stops at the pre-existing `Fairstream_MoveSameSizeCollisionPrompts` `S_FALSE` assertion; an exact RED-tree control reproduces the same `ERROR_IO_INCOMPLETE`, so it is not an R1a regression and is not counted as qualification evidence.
