# E7 interrupted-Move notice identity evidence

## Identity and result

- Scenario: test-enabled x64 Debug focused Commands case
  `cmd_fileops_move_breadcrumb_durable_lifecycle`.
- Machine profile: `4cb089111a23`.
- Runner run id: `20260831T102904Z-96656-d90fe2ef8fba473ea5402d5983d80c40`.
- Exact implementation commit: `5e6213e9b9cfbb4ebbbefcdde0354dd015b75344`.
- Workspace snapshot: `08dd595a7158e008a2e27e07172305681a762b4fa2c3ce19873515240ddaa36a`.
- Build receipt: `faeb962ea72b53e29a061975737ab348d96d2d87e715641bdc845f9b2911966f`.
- Build diagnostics: 0 warnings, 0 errors.
- Result: 1 passed, 0 failed, 0 skipped; no flaky, quarantine,
  regression, or isolation classification.
- Test-sandbox disk audit: clean, 0 issues.

## Contract exercised

- Keeps the in-session card/action key distinct from the explicitly named
  historical interrupted-operation ID.
- Loads one and multiple valid breadcrumbs while rejecting malformed, oversized,
  directory, reparse, and noncanonical provider-instance records independently.
- Decodes comparison-only provider instance identity into pane navigation context
  and opens source and destination in their recorded pane/plugin roles.
- Expands the real auto-collapsed card and proves navigation plus exact dismissal
  are available without Resume, retry, rollback, cleanup, Cut-again, fabricated
  diagnostics, item reveal, or provider mutation.
- Dismisses one exact regular breadcrumb without retiring its valid sibling or
  touching source/destination objects.

## Curated files

- `commands_results.json` and `commands_trace.txt`: focused Commands status and trace.
- `selftest_results.json` and `selftest_trace.txt`: outer selftest status and diagnostics.
- `run_all_tests_results.json`: runner verdict, classifications, and clean disk audit.
- `workspace_final.json`: exact clean commit/workspace snapshot binding.

This is focused E7 package evidence, not a Fresh Full repository closeout.

