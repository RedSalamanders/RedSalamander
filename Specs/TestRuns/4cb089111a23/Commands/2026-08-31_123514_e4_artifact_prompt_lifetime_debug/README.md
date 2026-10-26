# E4 Batch Rename artifact-prompt lifetime evidence

## Identity and result

- Scenario: test-enabled x64 Debug focused Commands case
  `cmd_pane_batchRename_artifact_prompt_shutdown_during_nested_pump`.
- Machine profile: `4cb089111a23`.
- Runner run id: `20260831T103514Z-31592-a5fc24f532a848188db79dc6e826bb06`.
- Exact evidence commit: `217bff62bab70fb064f7574bd10af83dc932615c`.
- Implementation commit: `5e6213e9b9cfbb4ebbbefcdde0354dd015b75344`.
- Workspace snapshot: `73ec3957a8c79b5adddc958accb41bda81fd416e590bdda563417929465b1095`.
- Build receipt: `5eb4ec62c0e44f747709a3f290d805c3134992b3afe970b167e18668176fc9c9`.
- Build diagnostics: 0 warnings, 0 errors.
- Result: 1 passed, 0 failed, 0 skipped; no flaky, quarantine,
  regression, or isolation classification.
- Test-sandbox disk audit: clean, 0 issues.

## Contract exercised

- Drives the real nested HostShowPrompt message pump while FolderWindow teardown
  begins and proves FileOperationState destruction is deferred until prompt
  dispatch unwinds.
- Retains only taskId across the nested prompt, re-finds the task after the pump,
  and safely no-ops completion after shutdown.
- Proves a missing registered artifact-prompt payload completes the live task by
  taskId with shutdown status instead of parking its worker.
- Cancels before Batch Rename mutation and permits a fresh File Operations state
  after teardown.

## Curated files

- `commands_results.json` and `commands_trace.txt`: focused Commands status and trace.
- `selftest_results.json` and `selftest_trace.txt`: outer selftest status and diagnostics.
- `run_all_tests_results.json`: runner verdict, classifications, and clean disk audit.
- `workspace_final.json`: exact clean commit/workspace snapshot binding.

This is focused E4 corrective lifetime evidence, not a Fresh Full repository closeout.

