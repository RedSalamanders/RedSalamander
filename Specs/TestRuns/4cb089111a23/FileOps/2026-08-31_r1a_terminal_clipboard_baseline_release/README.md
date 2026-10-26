# R1a terminal and clipboard truth baseline

- Machine profile: `4cb089111a23`
- Source commit: `f40f83f6` (R1a RED receipt, before the production change)
- Build: test-enabled x64 Release; original full-solution receipt `286d6b8abae0b6aad771cb6d3fb5eccf5d39bc647618471a59fc204b27731269`
- Artifact recreation: the first external baseline directories were not retained. These archived samples were recreated from the same committed RED tree with project receipt `20d14ee14819b7c17bed213ae25389227935c81a1486cc913451027b79a8327e`; diagnostics were `0 warning(s), 0 error(s)`.
- Command: `RedSalamander.exe --fileops-selftest --selftest-case=Phase10_ClipboardAdmissionAndRetainedActions --selftest-fail-fast --selftest-no-activate`
- Scenario: five sequential independent processes; each initializes/finalizes 4,096 selected-root result slots and evaluates 1,024 no-provider clipboard-gate decisions
- `FileOps.SelfTest.R1a.TerminalFunnelUs`: `587`, `487`, `539`, `619`, and `594` us
- Nearest-rank p95: `619` us
- Resource shape: 4,096/4,096 builders/terminal stores and zero provider calls in every process; process-local fixture state is released at process exit

Every baseline process is intentionally RED (`passed=2, failed=1, skipped=0`). The legacy terminal sink rejects `0/1` duplicate stores, and all `1,024/1,024` failed-consumption decisions still permit mutation. The protected loop completes before the assertion that clipboard-consumption failure must terminate without mutation.
