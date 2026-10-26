# FolderView focus/selection Release baseline

This is the pre-resolver-change baseline for
`folderView_perf_focus_selection_state` on machine `4cb089111a23`.

- Result: 1 passed / 0 failed / 0 skipped.
- Build: test-enabled x64 Release.
- Build receipt: `d4857ae72bd280dfd07b6aa5e7ecd403d0e425d3fe0ab077a9f6b25f20eed317`.
- Snapshot commit: `87af5c2749dd9160ca34db21701d83f8094de14f` (base `8313c1950e69f7ef1e5e057c5b0850381f410d33` plus Phase 0 specs, observations, scenario, and the behavior-neutral Release test guard fix).
- Governed run id: `20260820T134407Z-101228-2c620b6b43944aac82c22e3ed0da606a`.
- Command: `RSBuildEnableTests=true; .\Tools\Run-AllTests.ps1 -Suite Commands -Configuration Release -CaseFilter folderView_perf_focus_selection_state -PerfBudgetPath Specs\Testing\FolderViewPerfBudgets.json5 -TimeoutMultiplier 8`.
- Environment: local console, 150% DPI, 120 Hz, NVIDIA GeForce RTX 5080; WARP and RDP matrix variants were not run for this baseline.

The legacy implementation produced 206 `folder.focus.resolve_us` rows and 206 accepted model-change observations. Its expected correctness diagnostics were two current-invariant violations and one focus-memory-cap violation; those are baseline facts, not accepted target behavior. Focus memory peaked at 601 entries and 103,400 owned UTF-16 payload bytes with zero evictions.

## Compact artifact manifest

The raw JSONL exceeded the checked-in archive limit. The original is retained locally under
`.build/PerfBaselines/FolderViewFocusSelection/2026-08-20_154800_release/` and is bound by the summary below.

| Artifact | Bytes | Rows | SHA-256 |
|---|---:|---:|---|
| Raw `perf_metrics.jsonl` | 3,138,197 | 8,459 | `d88a024b42749e14a9182e62e8e2aaacc38b3cd046495ed7cd343d957e624c1c` |
| Checked-in compact JSONL | 666,637 | 1,784 | `c5ed29b2e6d3d9c7415fc1b4f1e5734e9b7f094583856ee48d363f463d9b207d` |

Compaction preserves, in original order, every metric whose name starts with
`folder.` plus every `FolderView.ApplyCurrentSort` row. It removes unrelated
startup, icon, DirectWrite, rendering, and DxUi telemetry. The behavioral JSON
results and traces beside this README are copied unchanged from the same run.
