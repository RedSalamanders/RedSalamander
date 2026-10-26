# Observatory Track 6 Monitor bounded-pipeline checkpoint

- Date: 2026-07-16
- Machine hash: `4cb089111a23`
- Scenario: `monitor.chrome.dxui_toolbar_statusstrip`
- Result: passed
- Build: test-enabled x64 Debug Monitor (`RSBuildEnableTests=true`)
- Command: `.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --wait-instance --perf`

This focused checkpoint validates the Monitor chrome and frame metric contract after the Track 6 queue, retained
history, search, Clear, clipboard, Document snapshot, and worker-layout changes. The overload drill injected 80
events into a 24-event retained-history budget and completed with `queued=0`, `retained=24`, and `dropped=56`.
The normal 260-event batch also drained completely. Selection/caret reset and all required default frame metrics
passed.

Artifacts:

- `results.json`: structured checks and required-metric presence;
- `perf_metrics.jsonl`: captured Monitor queue, retention, search, batch, and frame metrics;
- `trace.txt`: concise execution trace.

This is checkpoint evidence, not the final Track 6 closeout archive. The deterministic active-ETW-consumer teardown
proof and focused ETW latency rerun remain required before Track 6 closes.
