# FolderView focus/selection review-remediation Release candidate

This is the repeated post-review candidate for
`folderView_perf_focus_selection_state` on machine `4cb089111a23`.

- Result: 3 passed / 0 failed / 0 skipped (`SelfTestRepeat=3`).
- Build: test-enabled x64 Release.
- Build receipt: `2e4baa0ad87984873b5116e8d7483df2cf0852f6fd6412a5668af62ee30e81b3`.
- Source snapshot: `9b2ec2b86516ececf9b3ddbca8ea1aa3011cad161dfe011d487ab3ca47da3a8e` (repository HEAD `8313c1950e69f7ef1e5e057c5b0850381f410d33` plus this operation's working-tree implementation/specification diff).
- Governed run id: `20260821T083548Z-80844-ad2e6191887b4374ac9b22cde8951cd0`.
- Command: `RSBuildEnableTests=true; .\Tools\Run-AllTests.ps1 -Suite Commands -Configuration Release -CaseFilter folderView_perf_focus_selection_state -PerfBudgetPath Specs\Testing\FolderViewPerfBudgets.json5 -TimeoutMultiplier 8 -SelfTestRepeat 3 -SkipBuild`.
- Environment: local console, 150% DPI, 120 Hz, NVIDIA GeForce RTX 5080; WARP and RDP matrix variants were not run.

All three repetitions kept the current-item and focus-memory caps valid. The
combined trace contains 634 `folder.focus.resolve_us` rows; p95 is 4,328 us,
versus 4,406 us for the accepted same-machine Release candidate (-1.8%, noise).
The preset comparison, using the documented one-sample override for this
correctness-oriented scenario's sparse frame rows, reported zero regressions,
three improvements, and two noise metrics. No hard budget was added.

Focus memory peaked at the 512-entry limit and 132,096 owned UTF-16 payload
bytes. The final cumulative eviction count was 1,293 because all three
repetitions intentionally reused the same pane-local bounded LRU.

## Compact artifact manifest

The mixed application JSONL exceeded the checked-in archive limit. The original
is retained locally under
`.build/PerfRawArchive/2026-08-21_103616_folder_view_focus_selection_review_remediation_release_repeat3/`
and is bound by the summary below.

| Artifact | Bytes | Rows | SHA-256 |
|---|---:|---:|---|
| Raw `perf_metrics.jsonl` | 11,691,107 | 31,864 | `c7c00130ac88dbbc75098d1a8e05d95dc7a30ce31ca4f50131a9124d35a46478` |
| Checked-in compact JSONL | 641,749 | 1,754 | `b7e7545455dff0843f50197c15d19934fb783166792a73dda2e829274b403356` |

Compaction preserves, in original order, every `folder.focus.*`,
`folder.focus_memory.*`, `folder.refresh.*`, and `folder.frame.*` row plus every
`FolderView.ApplyCurrentSort` row. It removes unrelated startup, icon,
DirectWrite, rendering, layout-detail, and DxUi telemetry. The behavioral JSON
results and traces beside this README are copied unchanged from the same run.
