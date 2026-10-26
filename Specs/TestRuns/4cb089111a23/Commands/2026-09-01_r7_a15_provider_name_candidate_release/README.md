# R7-A15 provider-name candidate

- Machine profile: `4cb089111a23`
- Exact implementation and test-synchronization commit: `e925018f5d40925d418563c7cbf0922a86d0145b`
- Build: test-enabled x64 Release; full-solution receipt `3d8be5dc3a2b7fad5119f3274d55aa94e626f9826e284ff1692c34ae723e4ae9`; diagnostics `0 warning(s), 0 error(s)`
- Debug control: test-enabled full-solution receipt `527d6dd1bcde11ecd2ab40a647448300935a2ae11f8fb42f3999ed169956164b`; diagnostics `0 warning(s), 0 error(s)`
- Command: `Run-AllTests.ps1 -Suite Commands -SkipBuild -Configuration Release -CaseFilter cmd_pane_batchRename_engine_large_preview_perf -TestRoot D:\RedSalamander.Perf -AllowAlternateVolumeTestRoot -RunId r7-a15-final-0N`
- Scenario: five sequential independent processes; each builds a 10,000-row preview through the deterministic typed provider name contract with two composite child-name queries per changed row and one parent listing
- `batchrename.preview.build_plan_us`: `80,564`, `81,362`, `79,758`, `81,030`, and `78,450` us; nearest-rank p95 `81,362` us
- `batchrename.preview.provider_name_validation.us`: `81,600`, `86,330`, `81,883`, `84,378`, and `83,969` us; nearest-rank p95 `86,330` us
- Candidate ceiling: `146,235` us (`max(1.50 * 96,235 us baseline p95, 96,235 us + 50,000 us)`); every sample is also below the `5,000,000` us absolute cap
- Invariants per sample: `20,000` provider-name queries, one parent listing, zero arena fallbacks, zero rejected rows, `3,100,000` canonical-key bytes, and `3,091,136` retained collision-index bytes
- Memory control: `cmd_pane_batchRename_collision_name_index_memory_gate` retained `36,700,160` bytes for 65,536 names of 240 UTF-16 units, below the `67,108,864`-byte limit

The unfiltered perf JSONL, exact Commands result, and trace for every process are retained in this archive. The same raw evidence is retained under `D:\RedSalamander.Perf\evidence\r7-a15-provider-name-candidate-release`.
