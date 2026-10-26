# E6 endpoint/parent artifact-candidate evidence

## Identity and result

- Scenario: test-enabled x64 Debug focused Commands cases
  `cmd_fileops_artifact_registry_claim_identity_join`,
  `cmd_fileops_artifact_candidate_index_scaling`, and
  `file_operations_phase0_typed_contract_source_guard`.
- Machine profile: `4cb089111a23`.
- Runner run id: `20260831T085156Z-101724-6322b4508f724b82a236d337fc958d4e`.
- Exact implementation commit: `48faff7e4dae2a9c04f41efa0b4e7fd0075b07ae`.
- Workspace snapshot: `4bd481755d364c66d0183a7423eae45d4abdffb1d95a84b44dd8feecb75a5c83`.
- Build receipt: `cb5f353da8be0661a993e4a4723b96305291d9f1fe10b84e306909ec18fbff18`.
- Build diagnostics: 0 warnings, 0 errors.
- Result: 3 passed, 0 failed, 0 skipped; no flaky, quarantine, regression, or
  isolation classification.
- Test-sandbox disk audit: clean, 0 issues.

## Contract exercised

- Replaced the process-wide leaf hint with one immutable index scoped to the
  exact plugin, canonical instance, path profile, root, identity domain, and
  registry generation.
- Proved exact provider-parent/leaf hits, unrelated endpoint and parent
  suppression, and conservative probing after registry-generation drift.
- Preserved Possible name-shape visibility and the claimed-path identity-
  mismatch downgrade; the index remains only a `worth probing` filter.
- Proved FolderView and Find use the scoped index and no production consumer
  retains `HasClaimLeafHint`.
- Kept exact claim path plus current no-follow provider identity as the sole
  Proven and recovery authority.

## Measurements

All lookup scenarios use 10,000 ordinary rows. The retained raw metric stream
contains 443 rows and is 166,588 bytes, so no compaction was required.

| Exact-domain claims | Folder lookup | Find lookup | Folder binds | Find binds |
|---:|---:|---:|---:|---:|
| 1 | 44,477 us | 61,825 us | 0 | 0 |
| 100 | 42,854 us | 61,394 us | 0 | 0 |
| 10,000 | 43,314 us | 61,545 us | 0 | 0 |

- The 10,000-claim index build took 88,269 us in test-enabled Debug.
- Lookup remained flat across the three claim counts and every ordinary-row
  bind count remained zero.
- These measurements characterize candidate lookup and bind avoidance. They do
  not grant artifact classification, mutation, cleanup, or recovery authority.

## Curated files

- `commands_results.json` and `commands_trace.txt`: focused Commands status and
  trace.
- `selftest_results.json` and `selftest_trace.txt`: outer selftest status and
  diagnostics.
- `run_all_tests_results.json`: runner verdict, classifications, and clean disk
  audit.
- `workspace_final.json`: exact clean commit/workspace snapshot binding.
- `perf/perf_metrics.jsonl`: complete 443-row metric stream in original order.

This is focused E6 package evidence, not a Fresh Full repository closeout.
