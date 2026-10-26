# Review and retirement of the five retained stashes

Date: 2026-10-19. Reviewed source: `30c6cce0b69895fea6a0c68e8eb3e6c58f6bcdb4`
(local and remote master at the start of this review).

## Completed review checklist

- [x] Inventory all five immutable IDs, base/index/untracked parents and ancestry.
- [x] Preserve every stash and its parents in a verified incremental recovery bundle.
- [x] Review all 156 tracked file changes and the one untracked completed plan.
- [x] Compare exact integration history and surviving code, tests and current specifications.
- [x] Determine whether useful changes remain missing: none found; all five stashes are obsolete or superseded.
- [x] Repair exact-path admission of the five recent File Operations closeout records, with matching tests, inventory, developer documentation and normative lifecycle guidance.
- [x] Complete focused validation and final review before publishing this record.

This is the completed content review and retirement decision. Publication and stash removal are
ordered mechanical follow-up actions: publish this record and its governance correction to remote
master, verify that publication, then drop only the five recorded stash IDs and remove their temporary
recovery refs. The final local Git-state receipt belongs beside the recovery bundle; this immutable
record does not claim that later Git actions have already happened.

## Recovery and scope

Raw patches, manifests, integration comparisons, check results and `stashes.bundle` are retained at
`D:\RedSalamander.Perf\evidence\stash-review-20261019` outside the source tree, within the marked
local evidence sandbox. The verified bundle SHA-256 is
`A7C8F5D6D57C57866733D87A8351ADF1A03959966B8EAED9B327EF73B4642CF1`.
Its five advertised refs are `refs/codex-recovery/stash-review-20261019/0` through `/4`.
`git bundle verify` succeeded; this is incremental recovery and requires the master history already
present in this repository. Index parents are empty; the Clearflow stash alone has an untracked parent.

For example, recover the Clearflow stash without applying it:

```powershell
git fetch 'D:/RedSalamander.Perf/evidence/stash-review-20261019/stashes.bundle' refs/codex-recovery/stash-review-20261019/2
git stash store -m 'Recovered Clearflow stash' FETCH_HEAD
```

No stash was applied or popped for this review. Comparisons used immutable recovery refs and separate
indices beneath the evidence directory. Whole-file blobs and per-file patch identities establish
historical integration, not current runtime qualification. Normalized added-line comparisons were
navigation aids, followed by semantic review of the current owners. Aggregate reverse-patch checks
failed on changed contexts; their diagnostics remain in `reverse-landing-checks.json` and are not
reported as successful tests or proof that functionality is missing. The human-managed scratch plan
was excluded from content analysis.

## Stash decisions

| Original ordinal | Immutable stash | Tracked changes | Decision |
| --- | --- | ---: | --- |
| 0 | `ed7d9a5b7c6bb2707e5c3c30361999f1208eab7a` | 15 | Retire: October File Operations safety; current implementation or completed history supersedes it. |
| 1 | `64c4e39160e4fe2fead707534185a14c71780a71` | 59 | Retire: July UIA, TSF, product focus and test updates; current implementation or completed history supersedes it. |
| 2 | `785c9f5f5a223a4f24cbc8a40f8bed1ec50c730b` | 77 | Retire: June Clearflow File Operations consistency; current implementation or completed history supersedes it. |
| 3 | `0fff9871ad69d8abc89a501d39789137934d3239` | 1 | Retire: ViewerSpace TurboTreemap checklist; current implementation or completed history supersedes it. |
| 4 | `e310e4e574da408d79d2ec8a89afd1dbc9de7b0c` | 4 | Retire: Preferences migration and Squad notes; current implementation or completed history supersedes it. |

The October safety stash's base is an ancestor of master. Four of its fifteen per-file deltas match
`02841dd4b62022a5c3d04f72efa18a3dea84a2bb` exactly. Its remaining changes survive in expanded
implementations or renamed tests/current specifications. Provider-domain-aware confirmation formatting
remains in `FileOperationConfirmation.cpp` and its Runtime caller; an additional notice parameter and
qualified ingress anchor extend the contract. The item-result reducer still fails a committed Delete
that retains the selected root, with Skip/cancellation precedence and retained-storage facts tested.
Failure-surface Enter requires a fresh press and release on the armed Cancel-all button; autorepeat,
focus loss, hide and teardown cannot trigger it. Later failure-frame recovery strengthens the old
implementation. Existing Commands and reducer tests preserve the stashed scenarios. Reopening the
completed I27 plan or restoring the old includes/test filenames would regress lifecycle ownership.

The July stash base is from old side history rather than an ancestor of master. That does not mean
its changes are absent: thirty-one of fifty-nine per-file deltas match the later integrated commit
`53e7e6641525a5a7d5f39abc868cefdfa9bdaed9` exactly. The restored DxUi source was inspected at the
actual consumer pin `f65e4a4b9311370b3e3070088e2a8abe29adcb27` (API revision 4), not an unrelated
sibling checkout. Current TSF activation captures previous association/focus, checks failures, retires
owned state before callbacks, and guards rollback/deactivation against reentrant replacement sessions.
Current DxUi tests cover host focus transfer, root replacement, removal during focus notification,
reentrant TSF callbacks, host focus loss and hidden hosts. Accessibility resolves captured snapshot
records/paths for selection and table headers, including selected offscreen rows. Restoring the old
vendored source would create competing owners and remove later lifetime safeguards.

Product modal close helpers clear DxUi focus before HWND reset. Issues-pane closure uses the canonical
active-folder-view focus restoration helper; restoring its old duplicated logic adds no behavior.
Current Commands tests obtain the process-owned issues HWND instead of class-only global lookup.
Archive/modal keyboard closure and focus probes survive; missing old lines were temporary trace
plumbing. Do not restore per-message environment-file tracing or old private UIA helpers. The deployment
test now passes its test flag through child-process `AdditionalEnvironment`, avoiding global mutation.
Frozen test counts were superseded by source-derived inventories. The WarpDrive completed plan retains
the substantive continuation text and historical failures; it is not newly qualified by this review.

The Clearflow stash has sixty-seven of seventy-seven per-file deltas identical to integration commit
`5bf852933923cd7337e9e3af275ab758a629cea8`. Current copy/move/delete capability admission,
contained cancellation, typed conflict action policy, primary/overflow action layout, provider callback
and progress contracts, search/selection guards and test-only contained Delete TOCTOU hooks retain the
useful work. Current menu hover suppression additionally drains pending hover messages and uses a
captured point with a test injection seam. Pinned DxUi retains Label hit-test passthrough, captured Button
press through MouseLeave, MenuBar interaction reset and stable visible-row traversal. Restoring old
popup layout/summary strings would overwrite later behavior. The Czech and Slovak satellite corrections
must stay: this stash put French text in those two resources. Current localized positional resources
and their source-contract tests supersede it. The untracked Clearflow Done plan is already byte-for-byte
identical to current master (Git blob `6b91890e3d6508e957e25213383fb847ca24961f`).

The ViewerSpace stash only adds an unchecked WIP checklist. Its completed TurboTreemap plan contains
the checklist with later phase evidence and explicit evidence-based deferrals. Reintroducing the old
unchecked copy would misstate progress. The oldest Preferences stash contains no product implementation:
it changes retired Squad histories/inbox and migration checkboxes. Current Preferences explicitly persists
hidden/system-file settings, uses hosted prompts for reset, and has page/root/focus wrappers; old notes
claiming these are missing are obsolete. The completed migration record is the current historical owner.
No stale Squad notes or WIP checklist is revived.

## Governance correction

Read-only specification inventory exposed `DONE_HISTORY_CHANGED` for five exact 2026-10-19 files:
`FileOperations_CardCountsAndThroughput`, `FileOperations_DocumentationPrivacy`,
`FileOperations_LastNoteWrapping`, `FileOperations_UIReviewFixes`, and `ScreenshotSyntheticData`.
The correction admits only those five existing Markdown paths plus this completed review. It changes
no older Done file, baseline, wildcard or directory admission. The focused fixture admits all six,
rejects all six date-shifted siblings and an extra review companion, and keeps existing Done history
protected. `Tools/tool-inventory.json`, `Tools/README.md`, `Specs/README.md` and its normative consistency
entry describe the same bounded behavior. Admission is not application qualification.

## Validation

Before the admission correction, 221 read-only checks passed with no failures or skips: 204 harness
source contracts, 10 localization contracts and 7 source-derived test inventory cases. These verify
current policy/coverage structure; they are not native GUI, storage or Full qualification.

Focused governance validation passed 102 cases with no failures or skips: 55 specification
architecture, 18 documentation drift, 17 tool inventory and 12 tooling governance cases. Together with
the earlier 221 checks, this is 323 focused cases; these overlap the broader run and are not added to it.
Tool inventory classified all 197 discovered files with zero blocking findings. Specification inventory
has zero Errors; its 16 missing-anchor and 43 missing-target Information findings remain historical
nonblocking findings. Its protected Done baseline and existing historical bytes are unchanged.

The governed affected run `20261019T120421Z-10128-f12448b94636446ea079583c77bce156` completed
with all 15 selected entries promoted, reporting **1,227 passed, 0 failed and 26 skipped**. The initial
full-solution build, contained targeted deployment rebuild, and final full-solution re-attestation each
reported zero warnings/errors. Final Debug x64 receipt:
`63ef007e0c583295cc798f843290db02e6817ac85e987be728506ac8279a1923`.
The 953 read-only and 13 build-toolchain Pester cases all passed without skips; the full checked-in
archive inventory also passed. The suite breakdown below uses the runner's normalized counts rather
than combining native assertions, standalone harness summaries and Pester cases into a new total.

| Selected suite | Classification | Passed | Failed | Skipped |
| --- | --- | ---: | ---: | ---: |
| ToolsPesterBuildToolchain | PASSED | 13 | 0 | 0 |
| VcpkgMergeSynthetic | PASSED | 1 | 0 | 0 |
| CrashHandlingTests | PASSED | 1 | 0 | 0 |
| LocalizationTests | PASSED | 1 | 0 | 0 |
| SettingsSchemaTests | PASSED | 1 | 0 | 0 |
| RedConfigureTests | PASSED | 1 | 0 | 0 |
| MonitorTest | PASSED | 1 | 0 | 0 |
| ViewerSqliteTests.Noninteractive | PASSED | 1 | 0 | 0 |
| PluginContractTests | PASSED | 1 | 0 | 0 |
| ViewerPETests.Noninteractive | PASSED | 1 | 0 | 0 |
| ToolsPesterTests | PASSED | 953 | 0 | 0 |
| ProductUiTests | PASSED | 1 | 0 | 0 |
| CompareDirectories | PASSED | 249 | 0 | 26 |
| FileSystemCurlTests | PASSED | 1 | 0 | 0 |
| PerformanceTests2 | PASSED | 1 | 0 | 0 |

The first `-SkipBuild` attempt stopped before execution with exit 4 because it lacked a compatible
byte-verified full-solution receipt. That blocked attempt is retained separately; the subsequent
ordinary affected run rebuilt and passed selected work. Affected and PrePush explanations were inspected.
PrePush could not verify PR equivalence and retained local obligations; no CI delegation is credited.
The wrapper exited 2 with repository **NOT_EVALUATED**, despite the selected run's successful terminal
status. Commands, FileOperations, PE interactive, SQLite interactive and Monitor ETW latency remain
pending foreground agreement. The 26 Compare skips retain their environment prerequisites, including
live journal currentness and a configured FTP profile. The sandbox disk audit remains unclean with
274 issues, matching the previously recorded count; this review neither cleans it nor treats it as a
repository pass.

Raw results, explanations, the blocked attempt and compact `validation-summary.json` remain beside the
recovery bundle. The canonical checkpoint is under `D:\RedSalamander.Perf\evidence\runs`, and native
artifacts are under `D:\RedSalamander.Perf\runs` for that run ID. Application code, the DxUi pin,
admission implementation/tests, normative guidance and inventory remained fixed throughout the run.
Only this review's completion/result text was finalized afterward.


No product source or dependency pin changes are needed. The broad affected fallback executed the
native/build work above without foreground input. It does not establish Full, Release, ARM64, ASan,
external assistive-technology or scenario-specific performance equivalence. Existing historical failed
tests and deferred platform obligations remain unchanged. GitHub CI/CodeQL on the starting master were
prevented from running jobs by account billing/spending restrictions; no remote pass is inferred.

## Per-file disposition

There are 156 tracked entries because the same file can occur in multiple stashes. Every row below
refers to the stashed path; renamed/external owners are identified explicitly. `Exact` means the
per-file delta matches the named integration commit above, with the surviving owner inspected during
review. Other rows name the reviewed replacement or obsolete material. Historical identity is not a
new execution result. All rows have retirement disposition.

### 0: October File Operations safety

| Stashed path | Current owner | Disposition/evidence |
| --- | --- | --- |
| `RedSalamander/FileOperationConfirmation.cpp` | `RedSalamander/FileOperationConfirmation.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/FileOperationConfirmation.h` | `RedSalamander/FileOperationConfirmation.h` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/FileOperationItemResultReducer.h` | `RedSalamander/FileOperationItemResultReducer.h` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/FolderWindow.FileOperations.Popup.cpp` | `RedSalamander/FolderWindow.FileOperations.Popup.cpp` | Current release-armed failure actions and typed popup layout supersede old rendering; later recovery retained. |
| `RedSalamander/FolderWindow.FileOperations.Popup.h` | `RedSalamander/FolderWindow.FileOperations.Popup.h` | Current release-armed failure actions and typed popup layout supersede old rendering; later recovery retained. |
| `RedSalamander/FolderWindow.FileOperations.State.Runtime.cpp` | `RedSalamander/FolderWindow.FileOperations.State.Runtime.cpp` | Qualified confirmation endpoints and appended notice extend provider-aware source formatting. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.FileOps.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.FileOps.cpp` | Renamed current Commands coverage retains focus/keyboard contracts; trace/private-helper scaffolding superseded. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.Runner.cpp` | Runner renamed; current includes/registrations retain scenarios; temporary tracing retired. |
| `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.ItemResultReducer.cpp` | `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.Tests.ItemResultReducer.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.Phases10_13.cpp` | `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.Tests.Phases10_13.cpp` | Renamed native phase coverage retains the safety/cancellation scenario. |
| `Specs/Core/Core_SharedHelpers.md` | `Specs/Core/Core_SharedHelpers.md` | Current owning specification retains the durable safety/identity/key-release contract with later extensions. |
| `Specs/FileSystem/FileSystem_FileOperations.md` | `Specs/FileSystem/FileSystem_FileOperations.md` | Current owning specification retains the durable safety/identity/key-release contract with later extensions. |
| `Specs/Plans/WIP/README.md` | `Specs/Plans/WIP/README.md` | Completed history/current WIP index supersedes old checklist and qualification status. |
| `Specs/Plans/WIP/UI_FileOperationsSafetyRemediation_2026-10-01.md` | `Specs/Plans/Done/UI_FileOperationsSafetyRemediation_2026-10-01.md` | Completed history/current WIP index supersedes old checklist and qualification status. |
| `Specs/UI/UI_FileOperationsPopup.md` | `Specs/UI/UI_FileOperationsPopup.md` | Current owning specification retains the durable safety/identity/key-release contract with later extensions. |

### 1: July UIA, TSF, product focus and test updates

| Stashed path | Current owner | Disposition/evidence |
| --- | --- | --- |
| `Common/DxUi/DxUi.Accessibility.cpp` | `pinned DxUi: src/Controls/DxUi.Accessibility.cpp` | External pinned DxUi owns snapshot/focus/TSF implementation and stronger reentrancy tests; no vendored copy. |
| `Common/DxUi/DxUi.ComboBox.cpp` | `pinned DxUi: src/Controls/DxUi.ComboBox.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `Common/DxUi/DxUi.Controls.cpp` | `pinned DxUi: src/Controls/DxUi.Controls.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `Common/DxUi/DxUi.Grid.cpp` | `pinned DxUi: src/Controls/DxUi.Grid.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `Common/DxUi/DxUi.NativeTextInput.cpp` | `pinned DxUi: src/Controls/DxUi.NativeTextInput.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `Common/DxUi/DxUi.Tree.cpp` | `pinned DxUi: src/Controls/DxUi.Tree.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `Common/DxUi/DxUi.WindowHost.cpp` | `pinned DxUi: src/Controls/DxUi.WindowHost.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `Common/DxUi/DxUi.cpp` | `pinned DxUi: src/Controls/DxUi.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `Common/DxUi/DxUi.h` | `pinned DxUi: include/DxUi/DxUi.h`; `pinned DxUi: src/Controls/DxUi.h` | Exact integration delta; current owner retains behavior or current successor. |
| `Plugins/FileSystemDummy/FileSystemDummy.cpp` | `Plugins/FileSystemDummy/FileSystemDummy.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `Plugins/FileSystemDummy/FileSystemDummy.h` | `Plugins/FileSystemDummy/FileSystemDummy.h` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/CompareDirectoriesWindow.Internal.h` | `RedSalamander/CompareDirectoriesWindow.Internal.h` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/CompareDirectoriesWindow.cpp` | `RedSalamander/CompareDirectoriesWindow.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/ConnectionManagerWindow.cpp` | `RedSalamander/ConnectionManagerWindow.cpp` | Current hosted-dialog focus/teardown code extends integrated changes. |
| `RedSalamander/FolderView.Icons.cpp` | `RedSalamander/FolderView.Icons.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/FolderView.Rendering.cpp` | `RedSalamander/FolderView.Rendering.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/FolderWindow.FileOperations.Dialog.cpp` | `RedSalamander/FolderWindow.FileOperations.Dialog.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/FolderWindow.FileOperations.IssuesPane.cpp` | `RedSalamander/FolderWindow.FileOperations.IssuesPane.cpp` | Canonical focus-restoration helper replaces duplicated close logic. |
| `RedSalamander/FolderWindow.FileOperations.Popup.cpp` | `RedSalamander/FolderWindow.FileOperations.Popup.cpp` | Current release-armed failure actions and typed popup layout supersede old rendering; later recovery retained. |
| `RedSalamander/FolderWindow.FileOperations.State.Diagnostics.Part.cpp` | `RedSalamander/FolderWindow.FileOperations.State.Diagnostics.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/FolderWindow.FileOperations.State.cpp` | `RedSalamander/FolderWindow.FileOperations.State.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp` | `RedSalamander/FolderWindow.FileSystem.Commands.cpp` | Current Commands owner retains modal behavior; per-message trace plumbing retired. |
| `RedSalamander/FolderWindow.FileSystem.cpp` | `RedSalamander/FolderWindow.FileSystem.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/FolderWindow.ItemProperties.cpp` | `RedSalamander/FolderWindow.ItemProperties.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/NavigationView.Menus.cpp` | `RedSalamander/NavigationView.Menus.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/Preferences.CompareDirectories.cpp` | `RedSalamander/Preferences.CompareDirectories.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/Preferences.CompareDirectories.h` | `RedSalamander/Preferences.CompareDirectories.h` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/Preferences.Dialog.cpp` | `RedSalamander/Preferences.Dialog.cpp` | Current hosted page/root/focus management and settings persistence supersede older dialog wiring. |
| `RedSalamander/Preferences.HotPaths.cpp` | `RedSalamander/Preferences.HotPaths.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/Preferences.HotPaths.h` | `RedSalamander/Preferences.HotPaths.h` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/Preferences.Keyboard.cpp` | `RedSalamander/Preferences.Keyboard.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/Preferences.Plugins.cpp` | `RedSalamander/Preferences.Plugins.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/Preferences.h` | `RedSalamander/Preferences.h` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.CompareOptions.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.CompareOptions.cpp` | Renamed current Commands coverage retains focus/keyboard contracts; trace/private-helper scaffolding superseded. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.Connections.cpp` | Renamed current Commands coverage retains focus/keyboard contracts; trace/private-helper scaffolding superseded. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.Dialogs.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.Dialogs.cpp` | Renamed current Commands coverage retains focus/keyboard contracts; trace/private-helper scaffolding superseded. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.FileOps.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.FileOps.cpp` | Renamed current Commands coverage retains focus/keyboard contracts; trace/private-helper scaffolding superseded. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.Navigation.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.Navigation.cpp` | Renamed current Commands coverage retains focus/keyboard contracts; trace/private-helper scaffolding superseded. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.PluginConfig.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.PluginConfig.cpp` | Renamed current Commands coverage retains focus/keyboard contracts; trace/private-helper scaffolding superseded. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.Preferences.ChromeAndPlugins.cpp` | Renamed current Commands coverage retains focus/keyboard contracts; trace/private-helper scaffolding superseded. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.Preferences.FileOpsCompareAndTree.cpp` | Renamed current Commands coverage retains focus/keyboard contracts; trace/private-helper scaffolding superseded. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.Preferences.HotPathsAndKeyboard.cpp` | Renamed current Commands coverage retains focus/keyboard contracts; trace/private-helper scaffolding superseded. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ThemesGeneralPanes.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.Preferences.ThemesGeneralPanes.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.Preferences.ViewersAndKeyboardLists.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.Search.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.Search.cpp` | Renamed current Commands coverage retains focus/keyboard contracts; trace/private-helper scaffolding superseded. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.Settings.cpp` | Renamed current Commands coverage retains focus/keyboard contracts; trace/private-helper scaffolding superseded. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.ShellCommands.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.ShellCommands.cpp` | Renamed current Commands coverage retains focus/keyboard contracts; trace/private-helper scaffolding superseded. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.Shortcuts.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.Shortcuts.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.ViewCommands.cpp` | Renamed current Commands coverage retains focus/keyboard contracts; trace/private-helper scaffolding superseded. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.Runner.cpp` | Runner renamed; current includes/registrations retain scenarios; temporary tracing retired. |
| `Specs/Plans/WIP/Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md` | `Specs/Plans/Done/Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md` | Completed history/current WIP index supersedes old checklist and qualification status. |
| `Specs/Testing/Testing_PerformanceValidation.md` | `Specs/Testing/Testing_PerformanceValidation.md` | Exact integration delta; current owner retains behavior or current successor. |
| `Specs/Testing/Testing_TestCoverage.md` | `Specs/Testing/Testing_TestCoverage.md` | Current source-derived inventory and scoped evidence contracts replace stale counts/qualification text. |
| `Tests/DxUiTests/DxUiTests.Accessibility.cpp` | `pinned DxUi: Tests/Controls/DxUi.Tests.Accessibility.cpp` | External pinned DxUi owns snapshot/focus/TSF implementation and stronger reentrancy tests; no vendored copy. |
| `Tests/DxUiTests/DxUiTests.NativeTextInput.cpp` | `pinned DxUi: Tests/Controls/DxUi.Tests.NativeTextInput.cpp` | External pinned DxUi owns snapshot/focus/TSF implementation and stronger reentrancy tests; no vendored copy. |
| `Tests/README.md` | `Tests/README.md` | Current source-derived inventory and scoped evidence contracts replace stale counts/qualification text. |
| `Tools/Tests/RedSalamanderPluginDeployment.Tests.ps1` | `Tools/Tests/RedSalamanderPluginDeployment.Tests.ps1` | Current source-derived plan/count contracts or child-process environment replace frozen counts/global mutation. |
| `Tools/Tests/RunAllTestsPlan.Tests.ps1` | `Tools/Tests/RunAllTestsPlan.Tests.ps1` | Current source-derived plan/count contracts or child-process environment replace frozen counts/global mutation. |
| `Tools/Tests/TestInventory.Tests.ps1` | `Tools/Tests/TestInventory.Tests.ps1` | Current source-derived plan/count contracts or child-process environment replace frozen counts/global mutation. |

### 2: June Clearflow File Operations consistency

| Stashed path | Current owner | Disposition/evidence |
| --- | --- | --- |
| `Common/DxUi/DxUi.Controls.cpp` | `pinned DxUi: src/Controls/DxUi.Controls.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `Common/DxUi/DxUi.Grid.cpp` | `pinned DxUi: src/Controls/DxUi.Grid.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `Common/DxUi/DxUi.WindowHost.cpp` | `pinned DxUi: src/Controls/DxUi.WindowHost.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `Common/DxUi/DxUi.h` | `pinned DxUi: include/DxUi/DxUi.h`; `pinned DxUi: src/Controls/DxUi.h` | Exact integration delta; current owner retains behavior or current successor. |
| `Common/WindowMessages.h` | `Common/WindowMessages.h` | Exact integration delta; current owner retains behavior or current successor. |
| `Plugins/FileSystem/FileSystem.FileOps.cpp` | `Plugins/FileSystem/FileSystem.FileOps.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `Plugins/FileSystem/FileSystem.Search.cpp` | `Plugins/FileSystem/FileSystem.Search.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `Plugins/FileSystemDummy/FileSystemDummy.cpp` | `Plugins/FileSystemDummy/FileSystemDummy.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `Plugins/FileSystemDummy/FileSystemDummy.h` | `Plugins/FileSystemDummy/FileSystemDummy.h` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/CommandDispatch.Debug.h` | `RedSalamander/CommandDispatch.Debug.h` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/CompareDirectoriesWindow.Internal.h` | `RedSalamander/CompareDirectoriesWindow.Internal.h` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/CompareDirectoriesWindow.cpp` | `RedSalamander/CompareDirectoriesWindow.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/ConnectionCredentialPromptDialog.cpp` | `RedSalamander/ConnectionCredentialPromptDialog.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/ConnectionCredentialPromptDialog.h` | `RedSalamander/ConnectionCredentialPromptDialog.h` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/ConnectionManagerWindow.cpp` | `RedSalamander/ConnectionManagerWindow.cpp` | Current hosted-dialog focus/teardown code extends integrated changes. |
| `RedSalamander/ConnectionManagerWindow.h` | `RedSalamander/ConnectionManagerWindow.h` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/FindFilesWindow.cpp` | `RedSalamander/FindFilesWindow.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/FolderView.Enumeration.cpp` | `RedSalamander/FolderView.Enumeration.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/FolderView.FileOps.cpp` | `RedSalamander/FolderView.FileOps.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/FolderView.Icons.cpp` | `RedSalamander/FolderView.Icons.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/FolderView.Layout.cpp` | `RedSalamander/FolderView.Layout.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/FolderView.Rendering.cpp` | `RedSalamander/FolderView.Rendering.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/FolderView.cpp` | `RedSalamander/FolderView.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/FolderView.h` | `RedSalamander/FolderView.h` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/FolderWindow.FileOperations.Popup.cpp` | `RedSalamander/FolderWindow.FileOperations.Popup.cpp` | Current release-armed failure actions and typed popup layout supersede old rendering; later recovery retained. |
| `RedSalamander/FolderWindow.FileOperations.Popup.h` | `RedSalamander/FolderWindow.FileOperations.Popup.h` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/FolderWindow.FileOperations.State.Queue.Part.cpp` | `RedSalamander/FolderWindow.FileOperations.State.Queue.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/FolderWindow.FileOperations.State.Runtime.Part.cpp` | `RedSalamander/FolderWindow.FileOperations.State.Runtime.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/FolderWindow.FileOperations.State.cpp` | `RedSalamander/FolderWindow.FileOperations.State.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/FolderWindow.FileOperations.cpp` | `RedSalamander/FolderWindow.FileOperations.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/FolderWindow.FileOperationsInternal.h` | `RedSalamander/FolderWindow.FileOperationsInternal.h` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/FolderWindow.FileSystem.cpp` | `RedSalamander/FolderWindow.FileSystem.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/FolderWindow.h` | `RedSalamander/FolderWindow.h` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/Lang/cs-CZ/RedSalamander-cs-CZ.rc` | `RedSalamander/Lang/cs-CZ/RedSalamander-cs-CZ.rc` | Reject old French-in-satellite text; current Czech/Slovak localization is correct. |
| `RedSalamander/Lang/fr-FR/RedSalamander-fr-FR.rc` | `RedSalamander/Lang/fr-FR/RedSalamander-fr-FR.rc` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/Lang/ja-JP/RedSalamander-ja-JP.rc` | `RedSalamander/Lang/ja-JP/RedSalamander-ja-JP.rc` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/Lang/sk-SK/RedSalamander-sk-SK.rc` | `RedSalamander/Lang/sk-SK/RedSalamander-sk-SK.rc` | Reject old French-in-satellite text; current Czech/Slovak localization is correct. |
| `RedSalamander/NavigationView.Edit.cpp` | `RedSalamander/NavigationView.Edit.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/NavigationView.FullPathPopup.cpp` | `RedSalamander/NavigationView.FullPathPopup.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/NavigationView.Interaction.cpp` | `RedSalamander/NavigationView.Interaction.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/NavigationView.cpp` | `RedSalamander/NavigationView.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/NavigationView.h` | `RedSalamander/NavigationView.h` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/NavigationViewInternal.h` | `RedSalamander/NavigationViewInternal.h` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/Preferences.Dialog.cpp` | `RedSalamander/Preferences.Dialog.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/Preferences.FileActions.cpp` | `RedSalamander/Preferences.FileActions.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/Preferences.FileOperations.cpp` | `RedSalamander/Preferences.FileOperations.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/Preferences.Keyboard.cpp` | `RedSalamander/Preferences.Keyboard.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/Preferences.h` | `RedSalamander/Preferences.h` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/RedSalamander.cpp` | `RedSalamander/RedSalamander.cpp` | Current hover drain/point suppression and test seam supersede earlier menu wiring. |
| `RedSalamander/RedSalamander.rc` | `RedSalamander/RedSalamander.rc` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/Resource.h` | `RedSalamander/Resource.h` | Current resource allocation/summary strings supersede old numeric IDs. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.Connections.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.Dialogs.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.Dialogs.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.FileOps.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.FileOps.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.Preferences.ChromeAndPlugins.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.Preferences.FileOpsCompareAndTree.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.Preferences.ViewersAndKeyboardLists.cpp` | Renamed current Commands coverage retains focus/keyboard contracts; trace/private-helper scaffolding superseded. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.Preferences.cpp` | Renamed current Commands coverage retains focus/keyboard contracts; trace/private-helper scaffolding superseded. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.Search.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.Search.cpp` | Renamed current Commands coverage retains focus/keyboard contracts; trace/private-helper scaffolding superseded. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.Settings.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.ShellCommands.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.ShellCommands.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.Shortcuts.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.Shortcuts.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp` | `RedSalamander/SelfTest/Commands/Commands.Tests.ViewCommands.cpp` | Renamed current Commands coverage retains focus/keyboard contracts; trace/private-helper scaffolding superseded. |
| `RedSalamander/SelfTest/Commands/Commands.SelfTest.h` | Retired metadata; current Preferences implementation/completed plan | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/SelfTest/Common/SelfTestCommon.h` | Retired metadata; current Preferences implementation/completed plan | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp` | `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.Tests.Cases.SearchAndIndex.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.cpp` | Retired metadata; current Preferences implementation/completed plan | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.Phases05_06.cpp` | `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.Tests.Phases05_06.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.Phases07_09.cpp` | `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.Tests.Phases07_09.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.Phases10_13.cpp` | `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.Tests.Phases10_13.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.cpp` | Retired metadata; current Preferences implementation/completed plan | Exact integration delta; current owner retains behavior or current successor. |
| `RedSalamander/SplashScreen.cpp` | `RedSalamander/SplashScreen.cpp` | Exact integration delta; current owner retains behavior or current successor. |
| `Specs/Core/Core_CompareDirectories.md` | `Specs/Core/Core_CompareDirectories.md` | Exact integration delta; current owner retains behavior or current successor. |
| `Specs/FileSystem/FileSystem_FileOperations.md` | `Specs/FileSystem/FileSystem_FileOperations.md` | Exact integration delta; current owner retains behavior or current successor. |
| `Specs/Plans/WIP/Operation_Clearflow_FileOpsConsistencyParallelismSecurityPlan_2026-06-01.md` | `Specs/Plans/Done/Operation_Clearflow_FileOpsConsistencyParallelismSecurityPlan_2026-06-01.md` | Exact integration delta; current owner retains behavior or current successor. |
| `Specs/UI/UI_FolderView.md` | `Specs/UI/UI_FolderView.md` | Exact integration delta; current owner retains behavior or current successor. |
| `docs/FileOperations.md` | `docs/FileOperations.md` | Exact integration delta; current owner retains behavior or current successor. |
| Untracked `Specs/Plans/Done/Operation_Clearflow_FileOpsConsistencyParallelismSecurityPlan_2026-06-01.md` | Same path on master | Exact current Git blob `6b91890e3d6508e957e25213383fb847ca24961f`; already present. |

### 3: ViewerSpace TurboTreemap checklist

| Stashed path | Current owner | Disposition/evidence |
| --- | --- | --- |
| `Specs/Plans/WIP/ViewerSpace_TurboTreemap_SpeedAndMemory_2026-06-06.md` | `Specs/Plans/Done/ViewerSpace_TurboTreemap_SpeedAndMemory_2026-06-06.md` | Completed TurboTreemap checklist/evidence supersedes old unchecked WIP additions. |

### 4: Preferences migration and Squad notes

| Stashed path | Current owner | Disposition/evidence |
| --- | --- | --- |
| `.squad/agents/ripley/history.md` | Retired metadata; current Preferences implementation/completed plan | Retired Squad notes: obsolete bug claims/agent process history; current Preferences code and completed migration supersede. |
| `.squad/agents/sysadm/history.md` | Retired metadata; current Preferences implementation/completed plan | Retired Squad notes: obsolete bug claims/agent process history; current Preferences code and completed migration supersede. |
| `.squad/decisions/inbox/sysadm-f4-sync-unify.md` | Retired metadata; current Preferences implementation/completed plan | Retired Squad notes: obsolete bug claims/agent process history; current Preferences code and completed migration supersede. |
| `Specs/Plans/WIP/UI_PreferencesWin32Removal.md` | `Specs/Plans/Done/UI_PreferencesWin32Removal.md` | Completed migration replaces old unchecked F2 status. |
