# Terminal Kitty upload recovery, active close, and resize retention

- Status: passed
- Build: test-enabled Release x64 (`RSBuildEnableTests=true`)
- Scenario: `Terminal Kitty upload recovery active close and resize retention`
- Command: `.build\x64\Release\PluginContractTests.exe --terminal-selftests`
- Samples: five independent processes; every process passed `77/0`, exited 0,
  reached the Terminal unload quiet point, and wrote no stderr bytes.
- Upload recovery: each process used the real Ghostty-producing host lifecycle
  and a real Direct2D HWND render target. Eight upload rows covered a one-shot
  `ERROR_RETRY`, two differently scheduled transient images, stable
  `E_OUTOFMEMORY`, `D2DERR_RECREATE_TARGET`, and three successful publications.
  The final aggregate was one retry schedule, one stable failure, one device
  recovery, and two repaint requests per process. Successful real uploads were
  12-378 us.
- Active close: the worker entered a real 4096x4096 RGBA conversion before
  `Stop()` was requested. Joined close was 39,514-51,873 us (44,172-us median),
  and every process ended with zero queued, active-source, active-converted, and
  ready bytes plus no notification. This supported-machine result makes
  cooperative row/chunk stop polling unnecessary; the executable 5-second
  safety bound remains the cross-machine failure guard.
- Resize retention: each process performed 32 ordinary resizes while retaining
  the same one-pixel Kitty bitmap. Across 160 real render-target resizes,
  duration was 264-1,081 us (344.5-us median); resize-triggered recreation and
  reupload remained zero, with one retained bitmap/four retained bytes.
- Evidence quality: the five runs are same-machine, same-build, and
  content-free. The focused source-contract suite independently passed `30/0`.
- Raw evidence: `perf/perf_metrics.jsonl` contains 710 rows, is 338,024 bytes,
  and has SHA-256
  `eb44ec281e2d92f2b0352e6e615c706a963cc54f716a1808bf1c60c341a0891b`.
