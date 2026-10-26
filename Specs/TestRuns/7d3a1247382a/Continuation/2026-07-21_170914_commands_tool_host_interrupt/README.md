# Commands tool-host interruption continuation

## Context

- Date: 2026-07-21
- Branch: `codex/msdrive-mutation-reconciliation`
- Command: `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild`
- TestSandbox run ID: `20260721T150024Z-63056-ba7d0fabad60401b97e3bd7404b01877`
- Purpose: preserve the original Commands ordering after repairing `settings_store_search_roundtrip`.

The Codex tool host lost its output channel and exited. The repository's kill-on-close job terminated the test wrapper
and RedSalamander child. No RedSalamander crash dump or sidecar call stack was created; this was an orchestrator/tool
interruption, not an application crash.

The partial Commands result contains 170 recorded cases: 169 passed and 1 failed. The only failure is
`theme_v2_runtime_resolution_and_dynamic_perf`, whose maximum 380,873 us exceeded its 100,000 us guard. The repaired
`settings_store_search_roundtrip` passed in sequence. Execution stopped while entering
`cmd_pane_batchRename_engine_large_preview_perf`, before the Connection Manager and Preferences cases that motivated
the ordered rerun.

This archive is not an acceptance result. `abandoned-artifact-operation-owner.json` preserves the stale owner record;
per the shared artifact-lock contract, the next trusted test requires one serialized full-solution Debug rebuild.
