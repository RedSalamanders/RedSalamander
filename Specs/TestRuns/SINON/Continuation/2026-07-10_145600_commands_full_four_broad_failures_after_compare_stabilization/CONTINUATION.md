# Commands Full Broad-Order Failure - 2026-07-10 14:56 +02:00

## Command

```powershell
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2
```

## Result

- Run id: `20260710T123816Z-53544-f500bf1d9cde4344b52efcad609d02fb`
- Exit: `1`
- Summary: 782 passed / 4 failed / 2 skipped
- Disk audit issues: 0
- Fresh RedSalamander crash dumps: none observed.

## Failing Cases

- `cmd_preferences_dialog_category_tree_handles_reverse_keyboard_navigation`
- `cmd_pane_find_dialog_restores_reordered_sorted_grid_layout`
- `cmd_app_openDriveMenus_keeps_navigation_shell_stable`
- `cmd_app_swapPanes_keeps_navigation_shell_stable`

## Archived Evidence

- `wrapper.out.log`
- `wrapper.exit.txt`
- `run-all-tests-results.json`
- `selftest.results.json`
- `commands.results.json`
- `commands.trace.txt`

This is not accepted final evidence. The next continuation should root-cause and fix the Commands
broad-order failures directly; do not make them non-blocking or quarantine them.
