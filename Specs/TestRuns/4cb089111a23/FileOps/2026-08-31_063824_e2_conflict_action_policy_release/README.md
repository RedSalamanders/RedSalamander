# E2 conflict-action policy Release evidence

## Identity and result

- Scenario: test-enabled x64 Release
  `FileOpsFamily_Phase10_DeleteValidation` coverage for the E2 single-owner
  conflict-action policy and popup snapshot projection.
- Machine hash: `4cb089111a23`.
- Runner ID: `20260831T043801Z-77340-817830570189429ebc9fca78bbd7c276`.
- Result: `9/9` passed, `0` failed, `0` skipped in `6,279 ms`; runner wall
  time `11.9 s`; disk audit clean.
- Test-enabled Release build receipt:
  `ca505fa096ab5fa39a75458f5f7011809a6ed4ccc5f1d51f8b35d92ef4a020e6`.
- Source snapshot:
  `82503628fe4df99043ea25c18a815ed89c92bd2489e5f8e3dfff7d1be26b48f0`.
- The binary embeds repository HEAD
  `2942f44ad3ede90cddcf49bb76eb8a22eabe0ac5`; the receipt source snapshot
  additionally binds the uncommitted E2 implementation and tests represented by
  this archive.

The focused family covers permanent delete, deferred consent and exact Recycle
escalation, metadata/source-retention decisions, typed results, clipboard
admission, content verification, artifact touch guards, and the E2 conflict
policy matrix invoked at Phase 10 entry. Popup snapshot/default/Escape behavior
is additionally covered by the focused Commands runs recorded in the E2 plan
closeout.

## Candidate responsiveness measurements

The retained rows are the existing popup, conflict/consent, and DxUI frame
instrumentation emitted while the family ran. No new performance threshold is
introduced by E2.

| Metric | Rows | Average | p95 | Max |
|---|---:|---:|---:|---:|
| `FileOps.Popup.Render.TotalUs` | 76 | 37,257.8 us | 48,278 us | 50,404 us |
| `FileOps.Popup.HostedControls.SyncUs` | 76 | 879.7 us | 2,929 us | 4,179 us |
| `FileOps.Popup.Render.BuildSnapshotUs` | 76 | 74.3 us | 143 us | 233 us |
| `FileOps.Popup.WmPaintUs` | 76 | 38,610.4 us | 48,945 us | 51,364 us |
| `FileOps.Conflict.MetadataUs` | 23 | 5.8 us | 24 us | 83 us |
| `FileOps.Conflict.WaitUs` | 23 | 0 us | 0 us | 0 us |

There is no apples-to-apples checked-in pre-E2 Phase 10 popup stream: the E1
archives intentionally discarded unrelated popup rows, and the raw earlier
runner directories were removed by governed stale-run cleanup. Therefore these
numbers are candidate characterization, not an improvement claim. Correctness
comparison is supplied by the same-machine E1/E2 broad Release FileOps results
(`122/143`, zero failures in both), while numeric popup trend comparison remains
explicitly unbudgeted until a future repeated same-scenario baseline exists.

## Compact artifact manifest

The complete raw JSONL remains beneath the immutable local run directory at
`D:/RedSalamander.Perf/runs/20260831T043801Z-77340-817830570189429ebc9fca78bbd7c276/`.
It exceeded the per-file checked-in evidence limit. The compact stream preserves
the original row order and retains every metric whose name is `FileOps.Popup.*`,
`FileOps.Conflict.*`, `FileOps.Consent.*`, `dxui.frame.*`, or `DxUI::Paint`.

| Artifact | Bytes | Rows | SHA-256 |
|---|---:|---:|---|
| Raw `perf_metrics.jsonl` | 2,376,185 | 6,294 | `0822224ab322896e6b787efb25379640068caa1f267e4c7548597ed0ce142722` |
| Checked-in compact JSONL | 564,596 | 1,517 | `64e5dfef8575c97647d3c45b7db93b8baedbc2ab99866853f6e0b8cf858b1f0d` |

The JSON results and traces beside this README are copied unchanged from the
same runner-owned automatic archive.
