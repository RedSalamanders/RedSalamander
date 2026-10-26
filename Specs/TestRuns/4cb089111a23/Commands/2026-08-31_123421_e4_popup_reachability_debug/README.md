# E4 File Operations popup reachability evidence

## Identity and result

- Scenario: test-enabled x64 Debug focused Commands prefix
  `cmd_pane_fileops_conflict_`.
- Machine profile: `4cb089111a23`.
- Runner run id: `20260831T103421Z-41740-e1aa572e6ab143c187fcf5b85ebce740`.
- Exact evidence commit: `217bff62bab70fb064f7574bd10af83dc932615c`.
- Implementation commit: `5e6213e9b9cfbb4ebbbefcdde0354dd015b75344`.
- Workspace snapshot: `73ec3957a8c79b5adddc958accb41bda81fd416e590bdda563417929465b1095`.
- Build receipt: `5eb4ec62c0e44f747709a3f290d805c3134992b3afe970b167e18668176fc9c9`.
- Build diagnostics: 0 warnings, 0 errors.
- Result: 3 passed, 0 failed, 0 skipped; no flaky, quarantine,
  regression, or isolation classification.
- Test-sandbox disk audit: clean, 0 issues.

## Contract exercised

- Caption Close on the global File Operations popup remains hide-only and does
  not submit the focused task's decision or Cancel All.
- Publishing a newly actionable immutable conflict decision restores a popup
  hidden during metadata loading without submitting or changing the decision.
- An already-actionable decision remains parked after Close and is restored by
  the explicit Show File Operations command path.
- Escape on a visible actionable decision invokes that decision snapshot's
  published Cancel action, never the popup Close behavior or Cancel All.
- Conflict metadata still uses one provider roundtrip and the published action
  set retains its existing immutable ordering/default semantics.

## Curated files

- `commands_results.json` and `commands_trace.txt`: focused Commands status and trace.
- `selftest_results.json` and `selftest_trace.txt`: outer selftest status and diagnostics.
- `run_all_tests_results.json`: runner verdict, classifications, and clean disk audit.
- `workspace_final.json`: exact clean commit/workspace snapshot binding.

This is focused E4 corrective evidence, not a Fresh Full repository closeout.

