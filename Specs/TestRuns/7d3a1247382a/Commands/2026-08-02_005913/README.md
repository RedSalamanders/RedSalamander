# Embedded terminal first-slice lifecycle evidence

## Result

The focused x64 Debug Commands run passed
`cmd_pane_embedded_terminal_plugin_lifecycle` (1 passed, 0 failed, 0 skipped)
with runner id
`20260801T225910Z-37060-836f6a603a174f60a0db86bac8822729` and zero disk-audit
issues. The case opened a real ConPTY-backed `Terminal.dll` child in the
opposite pane, reached Running, reused the instance for both insertion modes,
and completed synchronous teardown.

## First-slice timing samples

| Metric | Sample | HRESULT |
|---|---:|---:|
| `terminal.open_to_visible_us` | 26,272 us | `S_OK` |
| `terminal.insert_context_path_us` | 328 us | `S_OK` |
| `terminal.insert_full_path_us` | 145 us | `S_OK` |

These are cold focused-run observations, not a statistical release baseline.
They prove that the named instrumentation is emitted and that the first-slice
operations complete successfully. Plan 005 still requires same-machine sample
sets, reviewed percentile budgets, Release measurements, and before/after
upgrade comparisons before any release performance claim.

## Reproduction

```powershell
$env:REDSALAMANDER_TEST_ROOT = 'D:\RedSalamander-terminal-gate0\.build\TerminalLifecycleTests2'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild `
  -CaseFilter cmd_pane_embedded_terminal_plugin_lifecycle `
  -TimeoutMultiplier 2
```

The compact raw evidence is retained beside this file in
`commands_results.json`, `commands_trace.txt`, `selftest_run_results.json`, and
`perf/perf_metrics.jsonl`.
