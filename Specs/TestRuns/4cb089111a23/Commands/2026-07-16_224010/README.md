# Observatory Track 5 — Make File List responsiveness and cancellation

- Date: 2026-07-16
- Machine profile: `4cb089111a23`
- Build: Debug x64, `master` at `f4e0c8c3bed86ed7ae8b1505b72d338b59e1ffd6` plus the working-tree Track 5 changes
- Command: `.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_makeFileList_generates_formats_and_saves_options --selftest-timeout-multiplier=2.0`
- Result: 1 passed, 0 failed, 0 skipped

The deterministic fixture contains 1,024 bulk files and 1,029 initial recursive entries. The JSON command returned
to the UI path in 15,444 us (test artifact; emitted metric 15,326 us), collected 1,029 entries in 44,542 us,
rendered 290,149 UTF-8 bytes in 24,447 us, transactionally published them in 3,873 us, and completed in 89,104 us.

The same case injects a 500 ms pre-collection delay, invokes the command again to request cancellation, waits for the
worker to stop, and verifies that no target was published. CSV and text generation also complete successfully.

Files:

- `commands_results.json` / `commands_trace.txt`: focused behavioral result and trace.
- `perf/make_file_list_metrics.json`: deterministic workload and UI-return summary.
- `perf/perf_metrics.jsonl`: command return, collect, generate, output, cancellation, feedback, and total metrics.

Related closeout evidence:

- Archive roundtrip/conflict behavior: `../2026-07-16_220811/`.
- Unpack prompt conflict policy: `../2026-07-16_220818/`.
- Final build log: `.build/logs/msbuild-20260716_223747_717.log` (0 warnings, 0 errors).
- Source contracts: 138 passed, 0 failed.
