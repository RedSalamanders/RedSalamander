# File Operations Popup Queue Reorder — Release

- Date: 2026-07-22
- Build: test-enabled x64 Release
- Source sandbox run: `20260722T030441Z-13384-e9bd026e0cb04a799c010feabbf45e05`
- Scenario: `Phase5_PreCalcCancelReleasesSlot`, repeated 10 times
- Result: 30 passed, 0 failed, 0 skipped (Setup + focused phase + Cleanup for every repeat)
- Command: `.\Tools\Run-AllTests.ps1 -SkipBuild -Suite FileOps -Configuration Release -CaseFilter 'Phase5_PreCalcCancelReleasesSlot' -SelfTestRepeat 10 -TimeoutMultiplier 2`

The focused phase covers adjacent queue Move Up/Down eligibility, engine order-key swapping, Start Now admission after reordering, cancellation, and slot release. All ten `FileOps.Queue.ReorderUs` samples are retained in `perf/perf_metrics.jsonl`.

| Metric | Samples | Min | Median | p95 | Max |
| --- | ---: | ---: | ---: | ---: | ---: |
| `FileOps.Queue.ReorderUs` | 10 | 23 us | 34 us | 41 us | 57 us |

This is a candidate-only measurement for a new feature, not an improvement claim. There is no pre-feature baseline for queue reorder because the operation did not exist.
