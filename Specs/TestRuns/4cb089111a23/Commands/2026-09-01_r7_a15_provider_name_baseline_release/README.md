# R7-A15 provider-name baseline

- Machine profile: `4cb089111a23`
- Production baseline commit: `ee6fdd875e94192d60e36fc7f2f1d04ed4da4ad3` (R7-A15 activation; the subsequently committed RED source guard was present in the test binary but does not change production behavior)
- Build: test-enabled x64 Release; full-solution receipt `421dc06995fd35877df5b2a9014767305ead8f2892ad8599c1fa1dff9e48031e`; diagnostics `0 warning(s), 0 error(s)`
- Command: `Run-AllTests.ps1 -Suite Commands -SkipBuild -Configuration Release -CaseFilter cmd_pane_batchRename_engine_large_preview_perf -TestRoot D:\RedSalamander.Perf -AllowAlternateVolumeTestRoot -RunId r7-a15-baseline-0N`
- Scenario: five sequential independent processes; each builds the existing Windows/path-identity Batch Rename preview for 10,000 deterministic rows
- `batchrename.preview.build_plan_us`: `90,647`, `96,235`, `90,951`, `91,039`, and `90,904` us
- Nearest-rank p95: `96,235` us
- Candidate ceiling: `146,235` us (`max(1.50 * baseline p95, baseline p95 + 50,000 us)`)
- Invariants: every sample reports 10,000 rows, 10,000 changed rows, zero error rows, and zero duplicate-fallback rows
- Memory control: `cmd_pane_batchRename_collision_name_index_memory_gate` retained `36,700,160` bytes for 65,536 names of 240 UTF-16 code units, below the `67,108,864`-byte limit

The unfiltered perf JSONL, exact Commands result, and trace for every process are retained in this archive. The same raw evidence is retained under `D:\RedSalamander.Perf\evidence\r7-a15-provider-name-baseline-release`.
