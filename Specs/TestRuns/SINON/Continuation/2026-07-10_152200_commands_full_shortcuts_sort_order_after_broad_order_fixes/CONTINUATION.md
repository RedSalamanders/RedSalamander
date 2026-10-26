# Commands Full Shortcuts Sort Failure - 2026-07-10 15:22 +02:00

## Command

```powershell
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2
```

## Result

- Run id: `20260710T130334Z-87100-0beadd9d6b674b5a9d56516967b41080`
- Exit: `1`
- Summary: 784 passed / 2 failed / 2 skipped
- Disk audit issues: 0
- Fresh RedSalamander crash dumps: none observed.

## Failing Cases

- `cmd_app_shortcuts_key_column_uses_natural_key_order`
- `cmd_app_shortcuts_restores_persisted_sort_order`

The previous four broad Commands failures were not present in this run. This is a new broad-order
Shortcuts sort-state blocker and is not accepted final evidence.

## Archived Evidence

- `wrapper.out.log`
- `wrapper.exit.txt`
- `run-all-tests-results.json`
- `selftest.results.json`
- `commands.results.json`
- `commands.trace.txt`
