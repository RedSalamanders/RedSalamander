# E4 Batch Rename worker-admission evidence

## Identity and result

- Scenario: test-enabled x64 Debug focused Commands case
  `cmd_pane_batchRename_admission_worker_owned_cancellable`.
- Machine profile: `4cb089111a23`.
- Runner run id: `20260831T081222Z-95224-fd9df41501e34f63999d77690a33648b`.
- Source base at capture: `c86bf581386da07524c051dbc9a2bc9f53ce24d1` plus the uncommitted E4 implementation under test.
- Source snapshot: `dcfc728351b894a1a7812216af874ef338080eb445e5d7f5813a2e9a93d496e9`.
- Build receipt: `28f959f6fe07017ce51e71d0961d1b1fb70b33944eb9bca92bff5b35fdd426b8`.
- Build diagnostics: 0 warnings, 0 errors.
- Result: 1 passed, 0 failed, 0 skipped; no flaky, quarantine, or isolation classification.
- Test-sandbox disk audit: clean, 0 issues.

## Contract exercised

- Characterized worker-owned qualification for 1, 64, and 1,024 rows against Local and a deterministic delayed provider.
- Proved the UI publication path performs no row-proportional capability, parent, or no-follow binding calls.
- Proved capability lookup, parent derivation, no-follow bindings, identity-domain validation, dependency/cycle construction, duplicate physical-object rejection, and immutable plan publication happen before mutation.
- Proved cancel-at-barrier, cancel-during-provider-call, and modeless-window close drain without mutation and reach the provider unload quiet point.
- Proved a hard-link alias is rejected and that navigation after qualification cannot retarget execution.

## Measurements

- Maximum complete UI publication path: 12,409 us (budget: 50,000 us).
- Maximum UI-only immutable capture: 707 us.
- Maximum end-to-end qualification wall time: 787,572 us.
- Maximum worker-only qualification time: 777,820 us.
- Maximum cancel-request-to-terminal latency: 16,000 us.
- Maximum retained admission bindings: 0.
- The 1,024-row scenarios emitted exactly 1,024 capability calls and at least 1,024 no-follow source bindings, all on the worker.

## Curated files

- `commands_results.json` and `commands_trace.txt`: focused suite status and trace.
- `selftest_results.json` and `selftest_trace.txt`: outer selftest status and diagnostics.
- `run_all_tests_results.json`: runner result, classifications, and clean disk audit.
- `perf/perf_metrics.jsonl`: the 92 claim-bearing Batch Rename admission rows retained in original order.
- `perf/compaction.json`: digest-bound provenance for the 6,632-row raw metric stream and the deterministic filtered stream.

This is focused E4 package evidence, not a Fresh Full repository closeout.
