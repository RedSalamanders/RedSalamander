# Observatory Track 6 Monitor closeout

- Date: 2026-07-17
- Machine hash: `4cb089111a23`
- Scenario: `monitor.chrome.dxui_toolbar_statusstrip` with the ETW latency burst
- Result: passed
- Build: test-enabled x64 Debug Monitor (`RSBuildEnableTests=true`)
- Command: `.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --wait-instance --perf --monitor-etw-burst-mode=latency --monitor-etw-burst-count=60 --monitor-etw-burst-size=260`

Correctness checks passed, including Clear selection/caret reset, bounded overload (`queued=0`, `retained=24`,
`dropped=56`), all 60 latency samples, complete queue drain, default frame metrics, and the ETW burst metric set.

Same-machine candidate metrics:

| Metric | p50 | p95 | p99 | max |
|---|---:|---:|---:|---:|
| `monitor.frame.append_to_visible_us` | 10,839 us | 28,086 us | 37,236 us | 37,236 us |
| `monitor.etw.batch_drain_us` | 764 us | 2,109 us | 6,353 us | 6,353 us |
| `monitor.frame.total_us` | 8,882 us | 22,189 us | 45,681 us | 158,998 us |
| `monitor.frame.present_us` | 148 us | 5,149 us | 12,858 us | 44,062 us |

The scheduling gates pass: batch-drain p95 is below 8,333 us, batch-drain p99 is below 16,667 us, and
append-to-visible p95 is below 50,000 us. The run also captured queue high-water, dropped-count, retained-line,
retained-byte, and incremental-search metrics.

Companion validation:

- MonitorTest Debug build `msbuild-20260717_080754_542.log`: 0 warnings / 0 errors;
- `MonitorTest.exe --document-model-selftest`: exit 0, including normal ETW join and deterministic 20 ms bounded
  fallback, WIL-owned Unicode clipboard storage, retention, and bounded file-reader coverage;
- test-enabled Monitor Debug build `msbuild-20260717_080805_119.log`: 0 warnings / 0 errors;
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`: 139 passed / 0 failed.

Artifacts:

- `results.json`: correctness checks, metric presence, and quantile summaries;
- `perf_metrics.jsonl`: complete captured metrics;
- `trace.txt`: execution trace.
