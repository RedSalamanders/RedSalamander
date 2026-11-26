# Pending-decision pointer and Start now desktop checks

The approved Release desktop window started at 20:36:25 CET. The launcher finished in
84.558 seconds and its owned app processes retired. It verified test-enabled x64 Release source
`cf21187a82ca3dcf1ac6755eee8fff027ee986f6` against receipt
`23fc8cee61ad83d109dc1e761094a59eb0a5f2aef23b848e7fef5340334be66e` before launch.

| Native case | Run | Passed / failed / skipped | Case time |
| --- | --- | --- | --- |
| `cmd_pane_fileops_conflict_two_pending_task_identity` | `20261126T193627Z-57132-2adbc481621c48b69539844146838186` | 1 / 0 / 0 | 2,477 ms |
| `Phase5_DiscoveryCancelReleasesSlot` | `20261126T193711Z-72412-38a1f155aea44733962c699b8e777e48` | 3 / 0 / 0 | 937 ms |

The Commands test checks the pending-decision disclosure's exact task, action, held focus and
unresolved choice. Before the click, it reveals B's disclosure again after the focus repaint
and requires the owned painted target. The Start now check invokes the popup command for one
queued task and verifies that it releases only the selected scheduling gate. The provider wait
and queued neighbor remain independently governed.

The tested Commands change and the pinned-formatted form have the same operations and assertions.
The C++ differences are confined to wrapping the new lookup call. The [popup test contract](../../../../UI/UI_FileOperationsPopup.md)
records the reveal and pointer checks.

Both exact runner selections exited 2 with their selected checks passing. The Commands native
result SHA-256 is `5e64cc9efadb979df749d30eb484bc9e92e8706662421e35032c1e26833ed9cd`.
The File Operations native result SHA-256 is
`c59c83d62332b64752c83e60354b938c67f2958da2111f90147a7c942fbec6ad`.
The native results are under `artifacts/selftest/last_run` in the completed-run archives.
The full desktop suites and hosted native platform gates remain open.

## Hosted x64 diagnostic

The [manual x64 Full run](https://github.com/RedSalamanders/RedSalamander/actions/runs/45841032537)
tested `27ee7c01991154fe5f662f1eada984146a768b36`. Its main native run
`20261126T162853Z-3084-677b79a846a94111acf8cd619ed03f4a` recorded 3,061 passes, one failure
and 69 skips. The pending-decision pointer and Start now cases both passed in the original broad
suites. Their case times were 3,506 ms and 860 ms.

The failed case was `FileOps_SelectionShapeLocal`. Its six concurrency scenarios passed. In
the later 4,097-file counting check, the totals, repeated callback, changed-metadata result,
source contents and destination-stage cleanup all met the assertions. The cancellation result
was `ERROR_PARTIAL_COPY`, where the fixture requires `ERROR_CANCELLED`. That failure is
tracked independently in the plan.

The downloaded artifacts are preserved in the governed import
`20261126T194318Z-72732-5a8d8276a4a0493f892f7d99cb64d5a4`. The hosted v1 report
SHA-256 is `6025f4ab1bea5db7613a229b20595a263e37ff1aec9b65b8eb5c5f13ee4d9716`.
