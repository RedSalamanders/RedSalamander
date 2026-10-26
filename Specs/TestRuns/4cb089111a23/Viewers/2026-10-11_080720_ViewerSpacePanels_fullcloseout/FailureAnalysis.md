# Final Fresh Full Release outcome

Run `20261011T053522Z-30852-a218b0a1e7c74f61a658ac29e58b2f37` executed all twenty entries with no reuse. Seventeen promoted; three failed. Repository verdict: **FAILED**. The report records 2,139 passed, 10 failed and 276 skipped cases. The reported suite duration is 3,595,967 ms. The solution Release build and post-toolchain rebuild completed without warnings or errors. Source and receipt identities are retained in the raw v2 report and receipt.

The unmodified v1 report preserves every declared case and skip reason. V2 retains source, build, plan, decision, promotion and verdict identities. `bindings.json` records original byte counts and SHA256. This is a complete Fresh run with a failed repository gate.

## Failed entries

- **ToolsPesterTests:** 928 passed, one failed. `keeps every approved post-Terminal Done history blob unchanged` rejects two unexpected, already committed documents: `Specs/Plans/Done/CodeReview_MasterMerges_2026-10-05/findings-ledger.md` and `Specs/Plans/Done/CodeReview_MasterMerges_Delivered_2026-10-10.md`. A read-only diagnostic against the test's base `6aecfde8e` confirms those paths, with no missing or modified baseline blobs. It is retained in `protected-history-diagnostic.json`; the history allowlist was not weakened.
- **Commands:** 937 passed, seven failed, six skipped. The seven failures below concern paths outside the changed ViewerSpace implementation. They also occurred in the earlier Full run; no retry classification evidence was collected, so they remain unclassified failures.
- **FileOperations:** two passed, two failed, 240 skipped after mandatory safety prerequisites failed. `FileOps_QueuedAdmissionContracts` lacks the refusal diagnostic for `trailing-dot.` (`0x8007007B`). `FileOps_SafetyRemediationContracts` failed its ordinary-child DACL-denied Copy witness: exact original security-descriptor/DACL/protection/control restoration failed. The harness retained the affected scratch tree and stopped dependent cases. `fileops-trace.txt` preserves the individual witnesses. Those 240 suite-stop skips are unaccepted qualification; no retained storage-proof scratch was manually cleaned.

| Commands case | Recorded failure |
| --- | --- |
| `file_operations_phase0_typed_contract_source_guard` | Clipboard admission source guard. Inspection finds a literal single-space assignment at `Commands.Tests.PluginConfig.cpp:6038`, while `FolderView.FileOps.cpp:874` uses column-aligned spacing and lines 876–877 install the accepted-admission callback. This identifies a stale formatting witness; the report remains failed. |
| `cmd_preferences_dialog_shell_footer_live_dx_interaction` | General did not refocus the menu-bar toggle before restoration. |
| `cmd_pane_find_dialog_result_shortcuts_use_shell_clipboard_and_file_actions` | Could not quiesce pre-existing File Operations before shortcut validation. |
| `cmd_compare_directories_options_uses_dxui_labels_without_visible_legacy_statics` | Reopened baseline probe did not find exactly one visible DxUi body host. |
| `cmd_pane_fileops_popup_throughput_disclosure` | Focused path tooltip was empty; expected 232 units for task 90001's source path. |
| `cmd_pane_fileops_conflict_two_pending_task_identity` | Task disclosure did not expose a measured pointer target. |
| `folderView_thumbnail_sort_popup_slider` | Real slider drag did not update the pane's thumbnail size. |

## ViewerSpace and other results

All seven noninteractive ViewerSpace cases passed. Its interactive test and repeated window churn passed, with owned jobs empty and exact saved foreground/keyboard focus observed at restoration. Commands records 32 passed ViewerSpace cases and six opt-in benchmark skips, including the corrected Explorer/Analyst composition and boundary capture case. The native and lifecycle frames use this Full run's receipt; screenshot and build hashes were independently checked. Debug checks have their separate exact receipts and scope.

ViewerPE's blocked-read Close test and the S3 multipart contract, which failed in the earlier Full run, both passed here. That does not establish a retry/flakiness classification for the earlier failures. Compare Directories passed with 30 conditional skips. PerformanceTests2 passed all fifteen tests. SQLite's interactive entry reports no compiled interactive cases; its successful exit does not imply interactive SQLite coverage.

No unrelated production code or failing assertion was changed to force the repository gate green. See [the raw dashboard](run-all-tests-dashboard.md), [v1 cases](run-all-tests-results-v1.json), [v2 provenance](run-all-tests-results.json), and [UI review](UIReview.md).
