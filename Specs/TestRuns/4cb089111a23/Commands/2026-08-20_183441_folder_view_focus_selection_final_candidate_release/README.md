# FolderView focus/selection final Release candidate

This is the final implemented focus/selection-state candidate for
`folderView_perf_focus_selection_state` on machine `4cb089111a23`.

- Result: 1 passed / 0 failed / 0 skipped.
- Build: test-enabled x64 Release.
- Build receipt: `d80e658f1df02cb5b42be5c82fa349f697118c9261ec579e31018bc56258ce03`.
- Source snapshot: `2f496c461d1d61509532f5f760015352c7d28a953af41e47d0ef8e30a7b70ae7` (repository HEAD `8313c1950e69f7ef1e5e057c5b0850381f410d33` plus this operation's working-tree implementation/specification diff).
- Governed run id: `20260820T163056Z-82244-416313433334403ab57fd61921519116`.
- Command: `RSBuildEnableTests=true; .\Tools\Run-AllTests.ps1 -Suite Commands -Configuration Release -CaseFilter folderView_perf_focus_selection_state -PerfBudgetPath Specs\Testing\FolderViewPerfBudgets.json5 -TimeoutMultiplier 8`.
- Environment: local console, 150% DPI, 120 Hz, NVIDIA GeForce RTX 5080; WARP and RDP matrix variants were not run.

The candidate produced 207 `folder.focus.resolve_us` rows for 208 explicitly
observed model changes inside the measured scenario interval. It completed with
zero current-invariant violations and zero focus-memory-cap violations. Focus
memory peaked at the 512-entry limit and 132,096 owned UTF-16 payload bytes,
with 89 deterministic LRU evictions.

The runner reported 17 disk-audit findings for pre-existing manually named
directories under the local `.build/TestSandbox/runs` root. They were outside
this runner-owned run, did not affect the 1/1 scenario verdict, and are retained
in `run-all-tests-results.json` as an environment caveat rather than rewritten
or deleted.

## Compact artifact manifest

The raw JSONL exceeded the checked-in archive limit. The original is retained
locally under
`.build/PerfCandidates/FolderViewFocusSelection/2026-08-20_183441_release/`
and is bound by the summary below.

| Artifact | Bytes | Rows | SHA-256 |
|---|---:|---:|---|
| Raw `perf_metrics.jsonl` | 3,896,576 | 10,628 | `7246a805e6750fdc37a8bc740ae41d9cccc08ca554a5de2fe9a37033b6fddd3d` |
| Checked-in compact JSONL | 674,247 | 1,824 | `44db12c9916675c4a965c13a4878a9975d2320bad2b6e97d24ebf40b0883ec95` |

Compaction preserves, in original order, every metric whose name starts with
`folder.` plus every `FolderView.ApplyCurrentSort` row. It removes unrelated
startup, icon, DirectWrite, rendering, and DxUi telemetry. The behavioral JSON
results and traces beside this README are copied unchanged from the same run.
