# Repository qualification: search, File Operations, S3 and harness follow-ups


## Current routing - 2026-10-22

Status: ACTIVE, I34. This separate non-normative record owns the cross-topic
investigations retained during the ViewerSpace consolidation. The original failures,
receipts, case matrices, hypotheses and reproduction requirements below are historical
observations from October 9-10, not newly reproduced failures at the current HEAD.
Source excerpts and old owner-rank references must be reconciled before execution.

The user asked to focus the [single ViewerSpace plan](ViewerSpace_ConsolidatedRoadmapAndQualification_2026-10-22.md)
on ViewerSpace and keep other topics in another spec. This record routes search/Compare,
File Operations/ACL/clipboard, S3, native viewer desktop restoration, link capability
and Commands prerequisites to their existing domain owners; it does not transfer those
implementations to ViewerSpace or create parallel domain ownership. Assign the accepting
person/agent before a repair. No new runtime, environment or repository verdict is claimed.

All original detailed observations are retained below. Resolve each family independently;
a later isolated pass does not close a historical broad-order failure. Preserve failed
artifacts and test/security/input guards. A recorded environment or CI failure requires
fresh evidence before being described as a current blocker.

## Investigation checklist

- [x] Preserve the exact Debug and Release run identities, receipts, aggregate counts and original failure evidence.
- [x] Record every failed Commands case, the native assertion witnesses and the five tooling cases.
- [x] Separate observed failures, candidate causes, setup/environment blockers and missing diagnostics.
- [x] Record the accepted ViewerSpace first-map targets and passing Space coverage separately.
- [ ] Assign a named maintainer to each family and reconcile the existing I29, I25, I27, I10, I4 and I2 ownership boundaries.
- [ ] Verify drift, source/build identity, test capability and the original fixture before reproducing a failure.
- [ ] Reproduce the selected family through the governed runner; include its preceding broad-process cases when isolation changes the outcome.
- [ ] Add missing boundary diagnostics and deterministic ordering witnesses before choosing a repair.
- [ ] Classify each finding as product defect, test defect, environment failure or inconclusive, with evidence and an owner.
- [ ] Repair confirmed defects in bounded owner-approved slices; preserve all original assertions and resource guards.
- [ ] Requalify the exact case, relevant predecessor cluster and broader owning suite; archive both failed and repaired results.
- [ ] Merge lasting behavior changes into the authoritative domain spec and close or route every finding before moving this investigation to Done.

## Status, scope and authority

- **Status:** ACTIVE investigation and routing record, I34. No application repair is implemented by this document.
- **Recorded:** 2026-10-10; moved out of the ViewerSpace backlog on 2026-10-22 with the original investigation evidence retained.
- **Priority:** P1 for cancellation, exact ACL restoration and desktop-restoration investigations; P2 for the other harness/UI/setup findings.
- **Owner:** the supervising broad-failure investigation maintainer/agent owns this record. Per-family implementation owners are roles below; people are not yet assigned.
- **Effort/risk:** M for initial characterization; repairs are estimated separately below. ACL/security and input ownership repairs require careful review.
- **Planned at:** commit `6c32c3d42af80e4b4d25b5ceedfb3b9a734847b9`, branch `codex/viewer-space-explorer-analyst-implementation`.
- **Category:** potential correctness bugs, test/harness defects and environment qualification.
- **Related delivery:** [ViewerSpace implementation and qualification](ViewerSpace_ConsolidatedRoadmapAndQualification_2026-10-22.md), [PR #118](https://github.com/DualTail/RedSalamander/pull/118). The user requested publication to remote master on 2026-10-10; the investigation remains ACTIVE after publication.
- **Authority:** this WIP is non-normative. Existing domain/testing contracts remain authoritative. Failure messages describe observations; they do not by themselves establish a product defect or attribute a failure to ViewerSpace.

This record persists the broader failures reported during ViewerSpace qualification. It owns evidence reconciliation, diagnosis and routing. It does not create a second owner for already active File Operations, S3, DxUi, tooling or FolderView implementation work. Coordinate a bounded repair with the existing owner before editing that domain.

Read these owning contracts before a corresponding investigation:

- [Local search](../../Core/Core_Search.md) and [Compare Directories](../../Core/Core_CompareDirectories.md).
- [File Operations](../../FileSystem/FileSystem_FileOperations.md) and [S3](../../FileSystem/FileSystem_S3.md).
- [Preferences](../../UI/UI_PreferencesDialog.md) and [FolderView](../../UI/UI_FolderView.md).
- [Viewer integration](../../Plugins/Plugins_ViewerPlugins.md).
- [Self-test execution and desktop ownership](../../Testing/Testing_SelfTests.md), [validation evidence](../../Testing/Testing_ValidationEvidence.md), [performance validation](../../Testing/Testing_PerformanceValidation.md) and [test guide](../../../Tests/README.md).
- Existing coordination: [I29 File Operations qualification](FileOperations_PostMergeQualification_2026-10-05.md), [I25 S3/harness residuals](FileOperations_S3PathStyleCopyAndHarnessResiduals_2026-09-15.md), [I27 DxUi qualification](DxUi_HostMessagesToolsetAndRestore_2026-10-01.md), [I4 source-contract migration](Operation_Astrolabe_TestContractArchitectureMigration_2026-07-17.md), [I10 tooling](Tooling_ToolsInventoryAndGovernance_2026-08-13.md) and [I2 FolderView qualification](Operation_FolderView_WarpDrive_RemainingCloseout_2026-08-25.md).

## Evidence identities and aggregate accounting

| Profile | Original run | Passed / failed / skipped | Build receipt |
|---|---|---|---|
| x64 Debug | `20261009T201507Z-45224-aa600fc03f9e4418b3b3e7152722bb0f` | 2,127 / 17 / 261; 2,405 total | `f721d4c3644b36016e3b81900978cc335f1a530d9eadc094f585683bc7f3638b` |
| x64 Release | `20261009T213050Z-43264-4bde93ea855e44c885a3277a68e89038` | 2,125 / 16 / 264; 2,405 total | `37dae646d3557dd0b5ac7618b76b136f0740f1be80dad2744b28236083f98ddd` |

Both aggregate change/repository verdicts are **FAILED**. Frozen build source is `9addaf53d4d88d618223a7f2d8554ac0b73f2472946c3a3e0875b1fd94ac9554`; the PrePush test snapshot is `430becd60c8af4bdcb56d4034e78e20028381748b4804249530bb9d0f3ee301e`. They are different identity domains. The root review verified ten frozen workspace files against both final test snapshots; the five tested code files remained byte-identical through subsequent documentation edits. `SOURCE_INPUT_CHANGED` is an impact-selection reason, not proof of mutation during execution. Receipts qualify the recorded source/runtime scope, not this later complete documentation snapshot.

Primary reports:

- [Debug aggregate v2](../../TestRuns/7d3a1247382a/Validation/space_debug_prepush_reports_20261009T201507Z/selftest/last_run/run-all-tests-results.json) and [case-level v1](../../TestRuns/7d3a1247382a/Validation/space_debug_prepush_reports_20261009T201507Z/selftest/last_run/run-all-tests-results-v1.json).
- [Release aggregate v2](../../TestRuns/7d3a1247382a/Validation/space_release_prepush_reports_20261009T213050Z/selftest/last_run/run-all-tests-results.json) and [case-level v1](../../TestRuns/7d3a1247382a/Validation/space_release_prepush_reports_20261009T213050Z/selftest/last_run/run-all-tests-results-v1.json).
- [Root qualification matrix](../../TestRuns/7d3a1247382a/Validation/space_finish_all_qualification_20261010/root-final-current-qualification-matrix.json).

| Aggregated suite | Debug failed | Release failed | Interpretation |
|---|---:|---:|---|
| Commands | 8 | 8 | Eight explicit case failures in each profile; the sets differ. |
| ToolsPesterTests | 5 | 5 | Five explicit symbolic-link/reparse tooling cases. |
| CompareDirectories | 1 | 1 | The same abort-match cancellation contract. |
| FileOperations | 1 | 1 | Safety-contract aggregate failed at the exact security-restoration witness. |
| ViewerPETests.Interactive | 1 | 1 | One failed native suite entry per profile; the native log contains additional containment/child assertions. |
| PluginContractTests | 1 | 0 | Debug aggregate exit failure includes FileSystem 149 passed/1 failed and S3 221 passed/3 failed. Release success does not exercise the Debug-only S3 assertions. |
| **Total** | **17** | **16** | These are aggregate reported failures, not 17/16 independent product bugs. |

ViewerSpace-specific Commands coverage was **34 passed / 0 failed / 4 declared opt-in skips in each profile**. Seven noninteractive Space native cases passed per profile; the separate exact Debug activation/Escape case passed. Release broad interaction reached and passed Space activation four times before later Web/Shell failures. The user accepted cached/small first-map p95 targets of **150/310 ms** on 2026-10-10; current measurements **145.603/302.094 ms** meet those named-fixture targets. The original archived 100/250-ms misses remain historical. None of these facts constitutes a green broader repository verdict.

## PB-01 — Parallel search loses abort-match cancellation

- **Observation:** `local_search_callback_contract` fails in both profiles. It expects `HRESULT_FROM_WIN32(ERROR_CANCELLED)` after the match callback aborts; the recorded return is `S_OK` (`0x00000000`). Debug duration 10 ms, Release 6 ms.
- **Fixture/expected behavior:** one `match.txt` containing `contract`; recursive forced scan, include files, wildcard `*.txt`, content search disabled. The same case separately checks success, progress failure, cancellation callbacks, cancellation and abort-progress. Only the final abort-match assertion is recorded as failing.
- **Impact if confirmed:** a caller may interpret an aborted search as a successful completion.
- **Confidence:** HIGH in the failed contract; MED in the candidate ordering mechanism. Suggested characterization S; repair M / medium risk.
- **Owner:** FileSystem search/Compare maintainer.

Source anchors: [assertion (line 81)](../../../RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.Tests.Cases.SearchAndIndex.cpp), [callback test implementation (line 3759)](../../../RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.Tests.Runner.cpp), and [parallel coordinator (line 934)](../../../Plugins/FileSystem/FileSystem.Search.cpp).

Observed current source excerpt:

~~~cpp
const HRESULT emitHr = EmitSearchMatch(runtime, match.metadata, match.matchedBy, match.contentResult);
if (FAILED(emitHr))
{
    RequestParallelDirectoryWalkStop(state, emitHr);
    break;
}
// Later in this iteration:
if (workersDone)
{
    break;
}
~~~

Candidate hypothesis: `terminalHr` and `workersDone` were sampled before result delivery. A failing callback sets shared terminal state, but an already-complete worker snapshot can break the loop before `finalHr` is updated. This is a source-level candidate; preserved metrics do not prove that ordering occurred.

Next steps:

1. Record callback invocation/result, shared terminal result, the worker-completion snapshot and the returned Search HRESULT for each subcall.
2. Force both orderings deterministically: callback abort after the last worker completes, and callback abort while work remains.
3. Reuse the existing cancel/error normalization and verify exactly one terminal outcome, worker retirement and no further result delivery after abort.
4. Do not change the expectation to accept success or introduce a sleep to hide the ordering.

[Detailed original triage](../../TestRuns/7d3a1247382a/Validation/space_finish_all_qualification_20261010/reproduction/space-current-prepush-compare-readonly-triage-20261009.md).

## PB-02 — ACL-denied Copy fixture cannot restore its original security state

- **Observation:** `FileOps_SafetyRemediationContracts` fails in both profiles. The Debug trace records `exact original SD/DACL/protection/control restoration failed; no source traversal or recursive cleanup`, then `ACL Copy cleanup FAILED to restore exact original owner/group/DACL descriptor/control; case retained.`
- **Expected behavior:** the ordinary-child DACL-denied serial Copy fixture must exercise its exact partial outcome and then restore the original owner, group, DACL, protection/control and source/destination trees before cleanup.
- **Impact:** failed fixture cleanup retains protected test data and prevents reliable subsequent qualification. The evidence does not yet distinguish a restoration API failure, representation mismatch or product/security-state change.
- **Confidence:** HIGH in restoration failure, LOW in cause. Characterization M; possible repair M / high risk.
- **Owner:** I29 File Operations qualification maintainer with the Windows security/test-harness owner.

Source: [restoration guard (line 456)](../../../RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.Tests.AclDeniedChild.cpp).
Evidence: [original Debug FileOps trace](../../TestRuns/7d3a1247382a/Validation/space_prepush_debug_failure_reports_20261009T201507Z/preserved-last-run/fileops/trace.txt). The native case remains retained after the failure; later gated/not-reached families are not successful coverage.

Next steps:

1. Capture the original and observed descriptor component/control differences plus restoration API result and Win32 error, scoped to the marked test-owned fixture.
2. Separate the Copy outcome witness from cleanup and identify the first state divergence.
3. Reproduce with the original ordinary-child denial, serial execution and exact tree checks.
4. Retain the fail-closed cleanup behavior. Never take ownership, grant blanket access, weaken exact restoration, follow links or recursively clean the retained fixture merely to make the runner green.

## PB-03 — S3 stalled-request cancellation proof does not reach its expected listing

- **Observation:** Debug PluginContract logs three failed assertions in the same bytes-flowing stalled-request scenario: recursive Delete must list the prefix; the request must return after Cancel; cancellation must return within a few body callbacks.
- **Additional witness:** `[S3] streaming request returned 15000 ms after Cancel`. This timing is diagnostic output, not a fourth failed assertion. Native S3 reports 221 passed / 3 failed; the encompassing PluginContract suite contributes one aggregate failure.
- **Expected behavior:** the fixture witnesses the listing, cancels the recursive operation, returns cancellation within the original bound and deletes no objects.
- **Confidence:** HIGH in the three recorded assertions, LOW in cause. Characterization M; repair M-L / high lifecycle risk.
- **Owner:** S3 provider/test maintainer; coordinate I25 and I29.

Source: [stalled-request proof (line 1503)](../../../Plugins/FileSystemS3/FileSystemS3.Tests.FakeS3.cpp), [prefix probing/listing (line 1505)](../../../Plugins/FileSystemS3/FileSystemS3.Directory.cpp), [request control (line 265)](../../../Plugins/FileSystemS3/FileSystemS3.Internal.h), [client creation (line 800)](../../../Plugins/FileSystemS3/FileSystemS3.Shared.cpp).
Evidence: [Debug PluginContract log](../../TestRuns/7d3a1247382a/Validation/space_prepush_debug_failure_reports_20261009T201507Z/checkpoint/entries/standalone.plugin-contract/results/PluginContractTests.output.log).

Preserve the current stimulus: 96-byte/40-ms dripping listing, initial parsed-GET polling window 10 seconds, 250-ms delay before Cancel, 15-second worker wait, less-than-3-second cancel assertion, configured connection/request timeout and provider watchdog. The fake recorded no parsed GET in its window. That does not prove bytes were actually arriving, which internal transport was used, or that a particular SDK control callback ran. The later endpoint stop occurs after the timing was sampled.

Next steps:

1. Record request method/path and timestamp, distinguishing a prefix probe from the recursive listing.
2. Locate admission, connection, response parsing, operation-control callback and worker-return boundaries.
3. Preserve the exact linked SDK/runtime revision; inspect that implementation before attributing failure to CRT, WinHTTP or another layer.
4. Cover bytes-flowing and stopped-body scenarios only when their own endpoint stimuli are witnessed; preserve cancellation HRESULT and no-DELETE checks.
5. Do not lengthen the waits, reduce the corpus or call Release aggregate success a qualification of Debug-only assertions.

[Detailed original triage](../../TestRuns/7d3a1247382a/Validation/space_finish_all_qualification_20261010/reproduction/space-current-prepush-s3-readonly-triage-20261009.md).

## PB-04 — Clipboard Move source guard is whitespace-sensitive

- **Case:** `file_operations_phase0_typed_contract_source_guard`, failed in both profiles with `Clipboard Move must be consumed at accepted admission, never deferred to terminal success.`
- **Observed source mismatch:** the guard requires `request.moveClipboardSequence = clipboardSequenceNumber;`; production source spells `request.moveClipboardSequence   = clipboardSequenceNumber;`. The source reader performs no whitespace normalization.
- **Impact:** valid admission code can be rejected by an exact formatting predicate. The failed source guard is not proof that runtime Clipboard Move consumption is broken.
- **Confidence:** HIGH in the current literal mismatch; the original process's resolved repository root was not recorded. Characterization/repair S / low-to-medium risk.
- **Owner:** Commands/FileOps test maintainer; coordinate I4 source-contract migration and I29 semantics.

Source: [literal predicate (line 6037)](../../../RedSalamander/SelfTest/Commands/Commands.Tests.PluginConfig.cpp), [production admission barrier (line 868)](../../../RedSalamander/FolderView.FileOps.cpp).

~~~cpp
// Guard:
folderViewFileOpsSource.find("request.moveClipboardSequence = clipboardSequenceNumber;")
// Production:
request.moveClipboardSequence   = clipboardSequenceNumber;
request.preWorkerReleaseBarrier = [clipboardOwner, clipboardSequenceNumber]() noexcept
{ return ConsumeMoveClipboardForAdmission(clipboardOwner, clipboardSequenceNumber); };
~~~

Next steps: record resolved absolute source root and file digest; independently evaluate all three guard predicates; prefer a behavior-level admission/consumption witness if the existing test surface supports it. A minimal formatting correction must preserve the check's actual semantic purpose. Verify accepted admission, rejected admission and pre-worker release; do not remove the guard or change Clipboard Move behavior based on this message alone.

[Detailed original triage](../../TestRuns/7d3a1247382a/Validation/space_finish_all_qualification_20261010/reproduction/space-current-prepush-clipboard-readonly-triage-20261009.md).

## PB-05 — Native viewer parent cannot prove exact desktop restoration

- **Debug:** `TestViewerWebUsesDxUiComboHostWithoutVisibleLegacyCombo` fails its parent foreground/keyboard-focus restoration witness. The suite stops before another child.
- **Release:** the same Web restoration witness fails during churn; `TestViewerShellComboHostsLongRunOpenCloseStayStable` then records failed fresh-harness/desktop-ownership restoration witnesses and the suite stops.
- **Expected behavior:** exact saved foreground HWND, saved GUI-thread keyboard focus, valid identity, input ownership and zero remaining owned child processes must be proved.
- **Confidence:** HIGH in failed restoration predicates, LOW in cause. Characterization M; repair M / high input-ownership risk.
- **Owner:** native viewer harness/DxUi maintainer; coordinate I27.

Source: [final restore predicate (line 1093)](../../../Tests/ViewerPETests/ViewerPE.Tests.Runner.cpp).
Evidence: [Debug native log](../../TestRuns/7d3a1247382a/Validation/space_prepush_debug_failure_reports_20261009T201507Z/checkpoint/entries/viewer-pe.interactive/results/ViewerPETests.Interactive.output.log) and [Release native log](../../TestRuns/7d3a1247382a/Validation/space_prepush_release_failure_reports_20261009T213050Z/checkpoint/entries/viewer-pe.interactive/results/ViewerPETests.Interactive.output.log).

The Debug Web child itself passed and retired its owned processes; saved identity/input and the immediate pre-restoration ownership checks passed. The final predicate failed, but the log lacks the activation API return and final foreground/focus values. It cannot identify which conjunct failed or prove user interference. The subsequent stop assertion is containment behavior, not another independent product cause. The Debug failure occurs before Space interaction in that suite; separate Space activation evidence remains valid.

Next steps: capture activation return, final foreground HWND/process/thread, `GetGUIThreadInfo` result/focus and saved identities at the failed boundary; establish an agreed focus-taking window; reproduce Web alone, its preceding PE combo case plus Web, then the owned Shell churn case. Preserve exact predicates, the directed warning, desktop lease, cursor/focus restoration and containment stop. Never attach foreign input queues, send focus/input to foreign windows or assume manual interference from the final boolean alone.

[Detailed original triage](../../TestRuns/7d3a1247382a/Validation/space_finish_all_qualification_20261010/reproduction/space-current-prepush-focus-restoration-readonly-triage-20261009.md).

## PB-06 — Link/reparse fixtures lack required environment qualification

Three recorded surfaces need separate classification:

1. Debug FileSystem native self-test: `relative file-symlink fixture should be created` fails; native result 149 passed / 1 failed.
2. Commands `file_action_selected_paths_identity_cleanup`: same-path reparse replacement/target should survive no-follow cleanup. Commands `file_action_selected_paths_recovery_bounds`: no-follow reparse fixture could not be created. Both fail in both profiles.
3. Five ToolsPesterTests cases fail in both profiles:
   - `rejects a final-file reparse destination without changing its external target`;
   - `rejects a real directory symbolic-link ancestor`;
   - `rejects a real file symbolic link and a dangling symbolic-link component`;
   - `refuses a test link that would live or point outside its marked sandbox, before creating it`;
   - `rejects reparse inputs and unsafe impact-base arguments without fetching`.

The tooling/native link evidence identifies a missing symbolic-link capability; it does not establish that every identity-cleanup assertion is explained by that capability. First prove fixture creation before assessing cleanup semantics.

- **Confidence:** HIGH in failure observations; setup classification is supported for the recorded privilege failures, while individual identity-cleanup causes remain unproven.
- **Owner:** Windows test-environment owner with Commands/FileSystem/tooling maintainers; coordinate I10.
- **Effort/risk:** S to inventory capabilities; any environment change belongs to its owner. Product cleanup repair, if confirmed, is M / high safety risk.

Source entry points: [Commands settings fixtures (line 3336)](../../../RedSalamander/SelfTest/Commands/Commands.Tests.Settings.cpp), [FileSystem binding self-test (line 9082)](../../../Plugins/FileSystem/FileSystem.FileOps.cpp). Exact Pester failures are retained in the case-level profile reports linked above.

Next steps: record effective capability, Developer Mode status, API return/last error and actual fixture identity without changing environment settings. Ask the environment owner for an explicitly authorized configuration when needed. Test-generated links and both their locations and targets remain inside the same owned marked sandbox. “External target” wording in a negative test is not permission to write or link outside that sandbox. Never silently skip the assertions, enable privilege/Developer Mode, or enumerate/clean unowned paths.

## PB-07 — Broader Commands UI state and prerequisite failures

This matrix includes all **ten unique failed Commands names** across the two profiles. Each profile has exactly eight failures. PB-04 and PB-06 own the first three rows; this section owns characterization of the other UI/prerequisite witnesses.

| Case | Debug | Release | Recorded failure |
|---|---|---|---|
| `file_action_selected_paths_identity_cleanup` | FAIL | FAIL | Same-path reparse replacement and target must survive no-follow identity cleanup. |
| `file_action_selected_paths_recovery_bounds` | FAIL | FAIL | Cannot create the no-follow reparse fixture. |
| `file_operations_phase0_typed_contract_source_guard` | FAIL | FAIL | Clipboard Move admission source guard. |
| `cmd_preferences_dialog_general_window_backdrop_apply_updates_supported_windows` | FAIL | — | Backdrop restore preview did not return to baseline theme state. |
| `cmd_preferences_dialog_mouse_focus_follows_pointer_settings` | FAIL | — | Apply did not rematerialize factory-default shortcuts in the running settings snapshot. |
| `cmd_preferences_dialog_shell_footer_live_dx_interaction` | — | FAIL | General page did not refocus the menu-bar toggle before restoration. |
| `cmd_pane_find_dialog_result_shortcuts_use_shell_clipboard_and_file_actions` | FAIL | FAIL | Pre-existing file operations did not quiesce before Find shortcut validation. |
| `cmd_compare_directories_options_uses_dxui_labels_without_visible_legacy_statics` | — | FAIL | Reopened baseline probe did not observe exactly one visible DxUi body host. |
| `cmd_pane_fileops_conflict_remote_case_sensitive_nested_inspection` | FAIL | FAIL | Both S3 panes did not enumerate the required real nested case twins. |
| `folderView_thumbnail_sort_popup_slider` | FAIL | FAIL | Real thumbnail-size slider drag did not update the pane through its size command. |

“—” means the case is not reported failed in that profile, not an inferred repair.

Source/registration anchors:

- [Preferences dispatch (line 287)](../../../RedSalamander/SelfTest/Commands/Commands.Tests.Preferences.Dispatch.cpp).
- [Find shortcut case (line 16460)](../../../RedSalamander/SelfTest/Commands/Commands.Tests.Search.cpp).
- [Compare-options case (line 5149)](../../../RedSalamander/SelfTest/Commands/Commands.Tests.CompareOptions.cpp).
- [Remote nested inspection (line 21839)](../../../RedSalamander/SelfTest/Commands/Commands.Tests.FileOps.cpp).
- [Thumbnail case (line 32911)](../../../RedSalamander/SelfTest/Commands/Commands.Tests.ViewCommands.cpp).

**Owner:** corresponding Preferences, Find/FileOps, Compare-options, S3 and FolderView Commands maintainers; coordinate I29, I27 and I2. **Confidence:** HIGH in recorded assertions, LOW in cause; characterization M overall, repair estimates require reproduction.

Each case requires its own prerequisite witness: exact settings/theme snapshot, expected focus, enumerations settled, prior FileOps quiet, current popup/host identity and actual slider input. Record broad test order and immediate predecessors. Run exact-case characterization first, then its predecessor cluster and broader Commands order. A standalone pass cannot close a broad-process failure. Do not presume all these failures share a single state leak, extend waits, accept duplicate hosts, or weaken the real-input assertions.

## Reproduction sequence and verified command surface

This document does not run new tests. A future executor starts from a native Windows checkout with the required toolchain and an initialized marked `X:\RedSalamander.Perf` root. `D:\RedSalamander.Perf` is the existing evidence-root example for this machine; an environment variable is a request, not initialization authority.

1. **Drift and owner check.** Run the following from the repository root and reconcile any changed source/test contract before proceeding:

~~~powershell
git diff --stat 6c32c3d42af80e4b4d25b5ceedfb3b9a734847b9..HEAD -- Plugins/FileSystem Plugins/FileSystemS3 RedSalamander/SelfTest RedSalamander/FolderView.FileOps.cpp Tests/ViewerPETests Tools/Tests Specs/Testing
.\Test-Changes.ps1 -Explain -Configuration Debug -TestRoot 'D:\RedSalamander.Perf'
~~~

The explain command produces the selected plan; it is not a runtime pass. Inspect declared interactive-desktop requirements, capability gaps, source identity and receipt obligations.

2. **Characterize a named noninteractive Compare case:**

~~~powershell
.\Test-Changes.ps1 -Suite Compare -CaseFilter local_search_callback_contract -Configuration Debug -TestRoot 'D:\RedSalamander.Perf'
~~~

Use test-enabled Debug and the original one-file query; obtain a new canonical run and preserve failed output if it still returns success.

3. **For a named Commands failure, use an exact case and the governed desktop lease only at an agreed time:**

~~~powershell
$failureCase = 'folderView_thumbnail_sort_popup_slider'
.\Test-Changes.ps1 -Suite Commands -CaseFilter $failureCase -Configuration Debug -TestRoot 'D:\RedSalamander.Perf' -Interactive
~~~

Substitute a case from the matrix only after inspecting its fixture and owning contract. The exact filter is essential; preserve focus restoration and the existing warning. Earlier session authorization is recorded historical evidence, not perpetual authorization for future focus-taking runs.

4. **Standalone native/provider/tooling cases:** obtain admitted entry IDs/arguments from `Tools/Modules/Testing/TestRunPlan.psm1` and the explained plan. PluginContract's noninteractive Debug gate is included in the governed PR surface:

~~~powershell
.\Tools\Run-AllTests.ps1 -Suite PR -Configuration Debug -TestRoot 'D:\RedSalamander.Perf'
~~~

This is a broader PR run, not a falsely advertised one-case S3 invocation. Do not launch an ad hoc DLL loader, invoke an internal child as a substitute for its parent harness or rely on process names for ownership.

5. **After a bounded repair:** follow affected iteration and the scoped PrePush policy in both required profiles. Interactive coverage requires an agreed time:

~~~powershell
.\Test-Changes.ps1 -Mode PrePush -Configuration Debug -TestRoot 'D:\RedSalamander.Perf' -Interactive
.\Test-Changes.ps1 -Mode PrePush -Configuration Release -TestRoot 'D:\RedSalamander.Perf' -Interactive
~~~

PrePush does not claim a Full repository pass. Required hosted/native/platform gates remain separate; CI last observed on head `bfbb76e2c` was unable to start steps due to account billing/spending-limit restrictions. Do not change billing or OS privileges as part of a code fix.

## Implementation boundaries and verification requirements

For the current request, edit only this WIP record and its index/backlink. Application fixes, tooling changes, environment changes and new tests are not implemented here.

A later activated slice must name exact files, owner, deterministic failing stimulus, expected terminal result and validation commands before source changes. Reuse canonical helpers from `Common/` and `Tests/TestSupport/`, WIL resource ownership, registered posted payloads and existing lifecycle/desktop-lease infrastructure. Native test additions follow the scoped naming/inventory policy. Tool changes require the tooling-governance workflow; they are not covered by this record alone.

Performance-sensitive fixes add/reuse boundary instrumentation, deterministic cases and equal-fixture archived before/after observations from the beginning. Capture app screenshots only through deterministic harness scenarios with provenance. Preserve all case counts, time/resource ceilings, cancellation expectations, complete detail and input/security predicates. Keep originals immutable; publish later qualification in new archives with exact run/source/build identities.

## Questions and evidence still needed

- Who is the named person/agent for each maintainer role, and which existing owner accepts each repair slice?
- Which coordinator ordering occurred when abort-match returned success?
- What exact descriptor component/control or API result caused ACL restoration to fail?
- Did the S3 operation reach admission, connection, parsed request or callback delivery before Cancel?
- Which final HWND/API/focus conjunct failed the native desktop-restoration check?
- Which Commands failures reproduce only after a specific predecessor, and which are independent?
- Which link assertions remain after an explicitly authorized environment capability is available?

These are diagnostic questions, not evidence that a repair requires guessing a new product policy. Continue read-only reconciliation and owner-independent investigation while missing runtime evidence is gathered.

## Done criteria and stop conditions

This investigation can close only when every PB family has a source-bound disposition and a named owner: repaired and requalified, confirmed environment blocker with explicit ownership, or separately accepted/routed residual with its own active record. Update this record and the accepting domain owner without claiming any unrun gate passed; ViewerSpace receives only a dependency disposition when relevant. Merge changed durable contracts into the authoritative specs, then move this completed record to `Specs/Plans/Done/` and update the index.

STOP the affected runtime/repair slice if its source/build/fixture differs without reconciliation, source ownership overlaps another active maintainer, focus time is not agreed, scratch/link paths cannot be proven sandboxed, cleanup cannot restore security exactly, cancellation/teardown ownership is uncertain, or validation requires weakening a guard. Preserve failed artifacts and report the precise blocker. Do not stop independent documentation/evidence reconciliation merely because one runtime slice is blocked.

Two failed reasonable repair attempts require a concrete diagnosis/owner review before another attempt; do not continue blind reruns or erase unsuccessful evidence. Hypotheses in this document remain hypotheses until their discriminating runtime witnesses are collected.
