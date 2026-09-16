# File Operations: accurate leaf progress and discovery scope

## Status and ownership

- Status: COMPLETE; C0-C4 implemented and verified on 2026-09-14. See the implementation record at the end of this plan.
- Index owner: I21.
- Priority: P1 for host discovery correctness; P2 for provider/API follow-ups.
- Effort: M for the leaf repair; L including nested-scope repair and qualification.
- Change risk: medium; closing discovery changes progress projection, discovery reservations, and verification totals.
- Planned at: `8aeb888e`, 2026-09-13. Drift check re-run at `63d78f55` on 2026-09-14: no changes in any owned path; every cited anchor re-verified against the source.
- Dependencies: no prerequisite implementation plan. Coordinate overlapping File Operations files with I18; I3 retains its existing stress-validation scope. Do not reopen retired I14/I16/I17 work.
- Authoritative contracts: `Specs/FileSystem/FileSystem_FileOperations.md`, `Specs/UI/UI_FileOperationsPopup.md`, `Specs/Plugins/Plugins_VirtualFileSystem.md`, and `Specs/Testing/Testing_PerformanceValidation.md`.

This plan is self-contained and non-normative. It addresses the reported single-file Copy/Move display and related discovery ownership errors. The audit was source-based; it did not run transfers, contact remote accounts/devices, or establish runtime reproduction for every provider route.

Run this drift check before implementation, then compare the excerpts below with current code:

```powershell
git diff --stat 8aeb888e..HEAD -- RedSalamander/FolderWindow.FileOperations* RedSalamander/SelfTest/FileOperations RedSalamander/SelfTest/Commands/Commands.SelfTest.FileOps.cpp Plugins/FileSystem Plugins/FileSystemCurl Plugins/FileSystemS3 Plugins/FileSystemGoogleDrive Plugins/FileSystemMicrosoftDrive Plugins/FileSystemMtp Plugins/FileSystemDummy Specs/FileSystem/FileSystem_FileOperations.md Specs/UI/UI_FileOperationsPopup.md Specs/Plugins/Plugins_VirtualFileSystem.md Specs/Testing/Testing_PerformanceValidation.md
git status --short
```

The unrelated `Dependencies/DxUi.lock.json` modification present at planning time was committed in `63d78f55`. Preserve all other work not owned by this plan. Reconcile changed implementation anchors before proceeding; a newer commit alone does not invalidate the plan.

## Problem and required outcome

The reported single-file Move displayed Running, approximately 33.8 MB/s, Discovering, Move: 0/?, and 424 MB of 424 MB so far. The host bridge transfers the file while leaving discovery open until its whole transfer call returns. It also omits the selected leaf's size from discovery accounting. The popup substitutes `max(discoveredTotalBytes, completedBytes)`, so the apparent total follows bytes already transferred.

A known regular leaf must publish its exact byte size and close its selected-root discovery scope after authoritative source binding/size acquisition and before payload transfer. Its task may remain Running, Verifying, waiting for a conflict decision, or completing source cleanup. Discovery closure is neither transfer completion nor permission to delete a source.

A directory must retain open discovery while its owning walker can discover more work. Nested provider calls must not close that directory's scope. Task discovery closes only when every selected root's discovery scope has resolved. A known empty file is zero bytes, not an unknown total.

## Prioritized findings

| ID | Finding and evidence at planned commit | Impact | Effort / risk / confidence | Disposition |
|---|---|---|---|---|
| D1 | `RedSalamander/FolderWindow.FileOperations.State.cpp:17069`: `TraverseAndPublishPath` counts a leaf, reports open discovery with no bytes, and closes on function exit after `CopyFile`. `:14379` already obtains the exact size in `BindSource`. | Single bridge leaf Copy/Move can spend its whole transfer showing discovery and no fixed progress denominator. | M / medium / high | Repair in C1. |
| D2 | Bridge options retain the root cookie at `State.cpp:11663`; `FinalizeManagedSourceCleanup:12048` passes those options into bound `DeleteIfUnchanged`. Local bound deletion emits `traversalClosed=TRUE` at `Plugins/FileSystem/FileSystem.cpp:3195` for bound files only (`_kind != FILESYSTEM_BOUND_DIRECTORY`); bound directory cleanup reports nothing. | During a Local-source managed directory Move, the first child file's cleanup can close parent discovery before the rest of the walk. Totals can subsequently grow after being presented as final. | M / medium / high, source-confirmed | Repair callback scope in C2. |
| D3 | Local `MoveItem` reports closed root discovery at `Plugins/FileSystem/FileSystem.FileOps.cpp:9394` before attempting rename. The host can continue a directory collision through `RenameMergeDirectory` using the same cookie at `State.cpp:18836` and `:18856`. Nested `RelocateChild` also calls provider Move with a copy of that root's options (`:15838`, `:15849`). Nested child reports do not corrupt totals: the walker publishes its cumulative counts before each nested call (`:16257` precedes `:16268`), so the cookie's monotonic guard zeroes the nested deltas. | Folder merge can start with discovery already latched closed; child rename callbacks can also prematurely close it. The defect is premature closure only, not count corruption. | M / medium / high, source-confirmed | Repair attempt and nested-call ownership in C2. |
| D4 | Curl, S3, Google Drive, Microsoft Drive, MTP, and Dummy have no production discovery reports on inspected mutation paths. Some already know leaf size; ordinary byte progress does not close host discovery. | Similar long-lived discovery display on admitted direct provider operations. Exposure varies by actual host strategy. | L across providers / medium / high for omission, medium for every route's popup exposure | Route-by-route follow-up in C4; do not rewrite all providers in the core patch. |
| D5 | Local `CopyItems:9741,9864` bypasses `ReportTopLevelDiscovery`; `MoveItems:10010,10134` reports individual roots through shared call options. Direct Delete reports open discovery without a close at `FileSystem.FileOps.cpp:4529,9559`. | Missing or non-cumulative discovery for direct bulk API consumers; late closure on direct leaf/recycle Delete. | M / medium / high | Separate API follow-up in C4. Current host rejects bulk Copy/Move. |
| D6 | `Phase8_PerItemOrchestration` asserts the progress denominator (`_progressTotalItems == 2`) at `SelfTest.Phases07_09.cpp:4953`, then bytes only after completion at `:5012`. Popup tests at `Commands.SelfTest.FileOps.cpp:4885` manually construct discovery state. | Existing tests can pass while live producer-to-popup discovery remains wrong. | M / low / high | Real active-task regressions in C0/C3. |

## Current implementation and invariants

The relevant single-root bridge code in `RedSalamander/FolderWindow.FileOperations.State.cpp:17100` is:

```cpp
else
{
    ++discoveredFiles;
}
const auto closeDiscoveryOnExit = wil::scope_exit([&]() noexcept { static_cast<void>(ReportDiscovery(0u, true)); });
HRESULT discoveryHr             = ReportDiscovery(0u, false);
// Link and directory branches intervene here.
const HRESULT fileHr = CopyFile(sourcePath, destinationPath);
```

`BindSource` at `:14376` has the necessary authoritative information:

```cpp
uint64_t& fileTotalBytes          = txn.fileTotalBytes;
bool& hasKnownFileTotalBytes      = txn.hasKnownFileTotalBytes;
const HRESULT hrReaderSize        = reader->GetSize(&fileTotalBytes);
if (SUCCEEDED(hrReaderSize))
{
    hasKnownFileTotalBytes = true;
}
```

The bridge's `discoveredBytes` counter is incremented only by the directory walkers (`:16254` sequential, `:16797` parallel). A selected leaf therefore contributes zero bytes to discovery even at the exit-guard close; `BindSource` feeds only the bridge's `totalBytes`, which reaches `_progressTotalBytes` through ordinary progress callbacks (`:7053`) and never the discovered total. `PumpAndPublishFile:14182` re-runs `PumpAndPublishFileOnce` on every replace re-prompt up to `kMaxReplaceRePrompts`, so `BindSource` executes once per attempt; a once-only leaf publication flag belongs on the traversal's leaf scope, not on the transaction.

`PumpAndPublishFileOnce:15453` executes `AdmitDestination`, `BindSource`, `RouteWriter`, `PrepareStage`, metadata, `Pump`, commit/proof, publication, verification, and cleanup in that order. Publish leaf discovery between successful source binding and writer/stage/payload work. Do not move source binding ahead of destination consent merely to improve the label.

`CopyFile:15514` passes `adoptFileSizeAsTotalWhenUnknown=true` for both a selected leaf and sequential directory children. That argument is **not** proof that the file owns the selected-root discovery scope. Use explicit scope information; do not infer root ownership from zero totals, one in-flight stream, path equality, or that existing argument. The leaf pump selects the buffered pipeline through `ShouldUseBufferedPipeline:12546`; fixtures force either mode through the existing `g_fileOpsBridgePipelineMode` override at `State.cpp:86` rather than a new switch.

`FileSystemReportDiscoveryProgress:7723` translates cumulative counts into deltas under `_discoveryMutex`, publishes progress, then calls `MarkDiscoveryItemClosed`. The latter's closure latch at `:7693` is one-way. `CloseDiscovery:7659` also releases discovery reservations and can freeze verification totals; verification totals are additionally raised per item at `:13248` through `AtomicMax`, so an earlier leaf close cannot understate them. Keep these consequences correct; resetting `_discoveryClosed` to repair an earlier mistaken close is not an acceptable design.

The provider ABI defines cumulative bytes/files/directories for the governed top-level call and monotonic counters within its cookie (`Specs/Plugins/Plugins_VirtualFileSystem.md:1170-1172`; cumulative counters are monotonic within one cookie at `:1176-1177`). A nested delete/rename is a different discovery scope even if its cancellation and bandwidth controls belong to the same task.

`Specs/FileSystem/FileSystem_FileOperations.md:620-624` already requires that the selected root be reported even when it is a single leaf, that a leaf's discovery record carry its exact file byte size and close that item's traversal before publication starts, and that recursive roots let their walker own root and descendant totals. The same rule governs the host bridge as traversal owner of bridge routes. D1 is therefore spec-versus-code drift, not a new contract: C1 implements the existing rule, and closeout spec work is limited to the nested-scope ownership rule for D2/D3 plus the popup and performance cross-references.

The popup's `ResolveWholeTaskProgressPresentation` at `FolderWindow.FileOperations.Popup.cpp:954` intentionally withholds percentages while `discoveryClosed` is false. `BuildTaskHeaderText:1316` similarly withholds the item denominator. The display at `:7340` uses discovered-so-far counters. Repair the producers and ownership; do not hide the bug by globally forcing determinate UI or treating any byte progress as proof of completed discovery.

Match existing WIL RAII, explicit HRESULT handling, atomics/locks, and typed task snapshots. Reuse `ReportDiscovery`, `MarkDiscoveryItemClosed`, and existing test seams. Search `Specs/Core/Core_SharedHelpers.md`, `Common/`, and `Tests/TestSupport/` before adding helpers. No raw owning handles, broad catches, new detached workers, or ungoverned posted payloads.

## Scope

Core implementation and test files:

- `RedSalamander/FolderWindow.FileOperations.State.cpp` and its existing `State.Private.h` / `FileOperationsInternal.h` support declarations; `State.Runtime.cpp` only if needed for existing debug-hook storage/reset.
- `RedSalamander/FolderWindow.FileOperations.State.Queue.cpp` for the Set/HasEntered/Release self-test wrappers of each new pause point (pattern at `:290-315`). A new `SelfTestPausePoint` also needs its extern in `State.Private.h:118-130`, its definition in `State.cpp:138-149`, and its wrapper declarations near `FileOperationsInternal.h:1990`.
- `RedSalamander/FolderWindow.FileOperations.Popup.cpp` / `.Popup.h` only for an essential snapshot/observation gap or an independently demonstrated projection defect.
- `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.cpp`, `.h`, and the relevant existing phase files; `RedSalamander/SelfTest/Commands/Commands.SelfTest.FileOps.cpp`.
- Existing plugin fake-backend/contract-test files only where needed to prove callback scope. Keep production provider parity work in the separately scoped C4 follow-ups.
- The four authoritative specs named above, this plan, and `Specs/Plans/WIP/README.md`; curated `Specs/TestRuns/` evidence during implementation closeout.

Do not change strategy qualification, source identity/cleanup authority, verification algorithms, link policy, admission/Preparing probes, concurrency limits, queue budgets, or public ABI layout. Do not add a totals pre-pass, recursive metadata scan for native rename, new setting, or hardcoded UI string. Do not edit `Tools/` for this repair; if a necessary tooling change emerges, apply tooling governance and scope it explicitly first.

## Execution checklist

### C0 — Characterize active discovery and protect existing behavior

- [ ] Verify route selection at `State.cpp:17916` and `:18729`: host bridge applies to managed Moves, separate destination providers, and qualifying verification-enabled Copy paths. Direct Local Copy is the positive control.
- [ ] Add registered cases named `DiscoveryScope_LeafCopy`, `DiscoveryScope_LeafMove`, `DiscoveryScope_ManagedDirectory`, `DiscoveryScope_NativeRenameMerge`, and `DiscoveryScope_MixedRoots` in the existing FileOps suite. Parameterize fixture variants instead of cloning tests.
- [ ] Register every `DiscoveryScope_` step in `kFileOpsFamilyDefinitions` (`FolderWindow.FileOperations.SelfTest.cpp:1852-1883`) under a new `FileOpsFamily_DiscoveryScope` family (or, if the owner prefers, inside `FileOpsFamily_Phase05_Discovery`). Unfiltered FileOps and Full runs enumerate family names only (`BuildRunFiltersImpl:2039-2049`); an unregistered step is silently skipped and would make this plan's Fresh Full gate vacuous. `-CaseFilter DiscoveryScope_` matches by prefix only because the trailing underscore selects prefix mode (`RedSalamander.cpp:7336`); the family filter is the Full-run guarantee.
- [ ] Reuse `SelfTestPausePoint` in `State.Private.h:8` for a bounded test-only checkpoint after successful source-size acquisition, before the first payload read. Observe the actual task and popup projection while the task is unfinished. Add a separate walker checkpoint after one nested mutation while siblings remain unresolved for D2/D3. The existing `g_fileOpsBridgeMoveSourceCleanupPausePoint` pauses before the nested delete (`State.cpp:954`), so it can serve only as the still-open control; the D2 observation needs a new pause point after the first child file's `DeleteIfUnchanged` returns (bound directory cleanup reports nothing), and the D3 observation needs one after the first child's native rename inside the merge walk. Wire each new pause point through the four files listed under Scope.
- [ ] Make every test release/reset its checkpoint on success, failure, cancellation, and teardown. The existing pause primitive has no task cancellation parameter; do not count time spent deliberately holding it as production cancel latency.
- [ ] Capture an expected-failing baseline for D1 and the nested-scope cases, plus passing direct Local Copy and wide-tree controls. Record exact strategy, endpoint topology, file size, configuration, and source revision.
- [ ] Run `FileOpsFamily_R4A19DiscoveryBaseline` and `FileOpsFamily_Fairstream` as controls: they hold the existing assertions on `firstMutationBeforeDiscoveryClosed` and closed totals (`Phases05_06.cpp:4350,4600,4769`, `Fairstream.cpp:4318`). All are recursive-tree fixtures, so C1 must leave them green; `R4A19_DiscoveryIndependentVolumes` additionally requires `-AllowAlternateVolumeTestRoot`.

Verify: build through the governed runner and run `-Suite FileOps -CaseFilter DiscoveryScope_`; the pre-fix result must fail on the intended active-state assertion, not a fixture setup or timeout. Existing controls listed under Commands below must pass or have their independent baseline failure recorded.

### C1 — Close selected regular-leaf discovery at authoritative size readiness

- [ ] Carry explicit selected-root versus descendant discovery responsibility through `TraverseAndPublishPath`, `CopyFile`, and the transaction boundary. Keep the directory walker the sole discovery owner for descendants.
- [ ] After successful `BindSource`, publish the selected regular leaf's exact size and existing one-file count through the canonical cumulative reporting path, with traversal closed, before writer/stage/payload work. Include zero bytes as a known total. Concretely: add `txn.fileTotalBytes` to the bridge's `discoveredBytes` exactly once per selected leaf and call `ReportDiscovery(0u, true)`; guard the once-only condition on the traversal's leaf scope because `PumpAndPublishFile:14182` can re-run `BindSource` on replace re-prompts.
- [ ] Update bridge discovery accounting itself, not only `_progressTotalBytes` or the popup. The exit guard must emit consistent already-reported totals and remain an idempotent terminal fallback.
- [ ] Preserve the existing source-size failure behavior: no invented zero-size success, no publication, and no destructive cleanup. Retrying the same scope must not double-add bytes/files or replay a closed scope as new discovery.
- [ ] Handle top-level reparse objects by semantic classification. `CopyLink:15709` routes bound regular-file placeholders into `CopyFile`; these need the same leaf fix. Bound directory placeholders at `:15719` retain actual traversal. Preserve/Skip links require item-based, no-follow accounting, not target enumeration or target-file byte totals.
- [ ] Confirm all selected roots must resolve before task closure. With several roots, earlier leaf transfers may correctly overlap unresolved root discovery. Do not pre-scan the selection to make the aggregate determinate sooner.

Verify: `DiscoveryScope_LeafCopy` and `DiscoveryScope_LeafMove` pass for both `g_fileOpsBridgePipelineMode` settings (buffered pipeline and single-buffer pump), positive/zero sizes, regular-file placeholders, verification on/off where admitted, and source-size failure/retry. At the held known-leaf checkpoint: task unfinished, one discovered file, exact bytes, discovery closed, no discovery activity, determinate byte or item progress. Payload completion and source retention/removal must remain correct.

### C2 — Isolate nested mutation discovery from the owning traversal

- [ ] Give nested cleanup, native child rename, and other bridge-internal provider calls an operation-control scope that forwards cancellation/deadline and discovery-mode requests, but cannot publish child-local totals or closure into the selected root. Audit every use of the bridge's `options`, including failure cleanup; apply the isolation consistently. Preferred vehicle: a second `PerItemCallbackCookie` (`FileOperationsInternal.h:913`) held by the bridge next to `options`, carrying the same `itemIndex` plus a new suppress-discovery flag that `Task::FileSystemReportDiscoveryProgress:7723` checks after header validation and before any counter update. Pass it as `operationControlCookie` for `FinalizeManagedSourceCleanup:12048`, `RelocateChild`'s `moveOptions` (`:15838`), and the failure-cleanup calls, while the `IFileSystemCallback` cookie stays the root cookie. `FileSystemShouldCancel:7353` and `FileSystemGetDiscoveryMode:7712` ignore the cookie, so cancellation and discovery-mode forwarding are unchanged. An `IFileSystemOperationControl` adapter is acceptable but not required.
- [ ] Preserve ordinary progress/conflict/result callbacks and their existing cookies. Do not replace cookies with an incompatible type accepted by `Task::FileSystemReportDiscoveryProgress`. Any scoped adapter must outlive the provider call and all objects opened under its options, per the existing ABI lifetime contract.
- [ ] Stage discovery closure from a provisional native directory rename until the host knows whether that attempt resolves the root or continues as a folder merge. A successful standalone native rename still completes as one root without enumerating descendants. A known directory-collision non-commit transfers discovery ownership to the merge walker without an earlier global close. Delivered mechanism, simpler than the one first sketched here: rather than suppressing the attempt's report and republishing counts, the host keeps the provider's exact cumulative counts and withholds only the one-way closure for the duration of the Native Move call. That needs no attribute pre-check, no size knowledge and no republication, and it is correct for a file rename too, because closure then comes from the per-item terminal guard at the same instant. On a directory-collision non-commit the merge bridge keeps the root cookie and closes at its true end.
- [ ] Keep the merge walk's cumulative accounting authoritative; do not double-count the initial root or child provider reports. A nested rename of an entire child directory does not authorize counting descendants that were never traversed.
- [ ] Retain one-way task closure, per-root idempotence, and existing terminal/cancel fallbacks. Do not reopen discovery or reset all cumulative counters after a collision/retry.

Verify: the managed-directory and native-merge cases hold after one nested operation while another child remains undiscovered and assert parent discovery remains open. At true traversal completion, closure occurs exactly once and totals stop growing. Include successful atomic directory rename, nested directory collision, Skip/Keep Both, cancellation, partial failure, and root order permutations. Use an existing initialized alternate fixed-volume fixture for an actual managed Local Move; do not substitute a same-volume native rename and label it managed.

### C3 — Verify the real popup and preserve streaming behavior

- [ ] Extend real active-task snapshot assertions for expanded/compact cards, item header, discovery indicator, progress meter, accessibility value, and footer/taskbar aggregate. Reuse the existing debug snapshot helpers instead of reading another process's windows.
- [ ] For a known nonempty leaf mid-transfer, require `0/1`, fixed full-file byte total, and determinate progress. Do not require an immediate ETA before the existing estimator has sufficient samples.
- [ ] For a zero-byte unfinished operation, require a valid item denominator and no divide-by-zero or false terminal success. Item completion still follows actual operation outcome.
- [ ] While one leaf task has closed totals and another directory task is still discovering, preserve the footer Known-work cohort and the taskbar's indeterminate rule.
- [ ] Preserve phase-5 wide-tree overlap, bounded queues, Skip's one-way switch to just-in-time discovery, Queue admission, pause/conflict presentation, cancellation, and teardown. Sequential directory traversal can legitimately remain open while processing its final entries.

Verify: all `DiscoveryScope_` cases, `cmd_pane_fileops_popup_progress_contracts`, phase-5 discovery, phase-8 orchestration, phase-11 bridge controls, the R4A19 discovery baseline family, and the Fairstream family pass. Tests must inspect real producer-to-snapshot state; synthetic snapshots and source-string assertions alone do not close D1-D3.

### C4 — Resolve and route adjacent provider/API findings

- [ ] For each provider row below, identify an actually admitted host route and record discovery events around its first real mutation/transfer using the existing fake backend. Existing `CancelControl` fixtures often discard discovery (`FakeS3.cpp:1405`, `FakeDrive.cpp:952`, `FakeGraph.cpp:853`, Curl `SelfTest.Ftp.cpp:4920`); extend observers rather than relying on final byte progress.
- [ ] Prioritize confirmed active leaf routes with size already available: Curl leaf Copy, S3 object Copy, and Dummy's delayed simulated transfer. Then qualify cloud-native Move/Copy and MTP routes against current capabilities. A missing callback is established; every claimed popup reproduction still needs host-route evidence.
- [ ] Write bounded successor WIP plan(s) for confirmed provider parity gaps, with exact producer, available size metadata, callback ownership, test fixture, and verification command. Use cumulative discovery for the governed call; preserve unknown-size semantics when metadata cannot prove a byte total. Never mark an unknown file as a known zero-byte file.
- [ ] Give Local direct bulk Copy/Move and direct Delete a separate API-contract successor, including serial/parallel cumulative accounting and one final call-level closure. Do not enable bulk Copy/Move in the host: `State.cpp:17343` intentionally rejects it before I/O.
- [ ] Before moving this core plan to Done, give each unresolved confirmed D4/D5 row a link to an indexed live successor owner, or record an explicit evidence-based rejection. Do not silently drop the findings or expand the core patch into all providers' recursive walkers.

Verify: every row below has a supported/unsupported route disposition and every accepted residual has a live indexed owner. Provider runtime claims require a passing named fixture; unexercised live environments retain explicit limitations. No unsupported provider feature is implemented as part of discovery repair.

## Adjacent-use-case audit and follow-up inventory

| Scenario | Current evidence and assessment | Required treatment |
|---|---|---|
| Cross-provider leaf Copy; managed cross-volume Move; verification-enabled leaf Copy | Shared bridge entry at `State.cpp:17916,18741`; D1 applies. | Core fix and route-asserting fixtures. |
| Several selected files; mixed files/directories | One closure slot per source (`State.cpp:7589,7693`). Leaf bug repeats, but aggregate cannot close before unresolved roots. | Core aggregation tests, concurrency 1 and greater than 1, both source orders. |
| Regular-file cloud/WOF/dedup placeholder | `CopyLink:15709` recognizes bound regular content and calls `CopyFile`. | Include in core leaf repair; Preserve/Skip must not mistake placeholders for semantic links. |
| Semantic link Preserve/Skip; directory placeholder | `CopyLink:15719,15734` separates directory namespaces and literal links. Outer scope currently closes on return. | Characterize item-scoped closure and no-follow behavior; never close a directory based only on the reparse attribute. |
| Managed directory Move from Local | Nested bound file cleanup reports closure through parent options (directory cleanup reports nothing); D2. | Core ownership repair and held-walk test. |
| Native folder Move becoming a merge | Initial and child native reports can close parent; D3. | Core attempt/child-scope repair. |
| Direct Local single Copy/native Move | `FileSystem.FileOps.cpp:4467,9243,9394` publishes root discovery before operation. | Positive control; no late-closure fix needed for standalone scope. |
| Sequential/parallel recursive Copy | Host sequential `State.cpp:16257` overlaps processing and traversal; parallel `:16960` closes before waiting for queued workers. Local equivalent at `FileSystem.FileOps.cpp:4938,5960`. | Intentional streaming controls; no second scan or blanket early closure. |
| Curl leaf Copy/native rename | `FileSystemCurl.CopyMove.cpp:2568,2611` knows leaf metadata; `Internal.h:723` supplies ordinary progress only. | D4 successor; fake FTP/SFTP-style route coverage as supported, no new transports. |
| S3 object Copy/Move entrypoints | `FileSystemS3.Directory.cpp:2661,2730` knows plan total before object work, without discovery reports. Host strategy may choose another path for Move. | D4 successor; qualify actual object Copy route first, preserve prefix semantics. |
| Google Drive server Copy/native Move | `FileSystemGoogleDrive.cpp:1800,1813` executes update/copy without discovery reports. | D4 successor using FakeDrive; distinguish files from recursive folder Copy and atomic folder Move. |
| Microsoft Drive native Move | `FileSystemMicrosoftDrive.cpp:6126` executes Move with terminal progress only. Direct Copy at `:5943` is unsupported. | Qualify Move in D4 successor; do not add direct Copy. |
| MTP Copy/Move entrypoints | `FileSystemMtp.Core.cpp:3616,3630,4408` has zero-byte start progress and backend work, no discovery reporter. Native Move remains a separately qualified capability. | Route/fixture characterization first; no capability promotion or live-device assumption. |
| Dummy delayed operations | `FileSystemDummy.cpp:6248,6283` mutates and simulates extended byte progress without discovery. | D4 successor and explicit discovery-aware fixtures; do not treat Dummy timing as physical transfer throughput. |
| Local direct bulk Copy/Move | Copy bypasses root helper; Move emits per-root closed totals through a shared call scope (`FileSystem.FileOps.cpp:9741,9864,10010,10134`). | D5 API successor, not the present popup's bulk path. |
| Local direct leaf/recycle Delete | Open reports at `FileSystem.FileOps.cpp:4529,7532,7571`; no close in `DeleteItem:9559`. Bound leaf Delete at `FileSystem.cpp:3195` already closes early. | D5 successor; distinguish direct/recycle, bound permanent Delete, and nested cleanup. Shell-owned recursive recycling may legitimately expose only selected roots. |
| 7z direct Copy/Move/Delete | `FileSystem7z.cpp:1506` onward returns `ERROR_NOT_SUPPORTED`. | Rejected as a direct mutation-discovery repair. Archive export through host bridge remains covered by bridge semantics. |

Audit exclusions: unrelated search/index discovery, Compare Directories scanning, archive creation internals, live remote credentials/devices, exhaustive provider recursion/performance, and unrelated security/dependency reviews. This is a discovery/progress audit, not a complete File Operations correctness audit.

## Commands and verification gates

Use repository-root PowerShell, the governed runner, and an already initialized marker-owned fixed-drive test root. Replace `X:\RedSalamander.Perf` below with that exact existing root; an environment variable alone is not ownership. First creation or alternate-volume setup needs explicit initialization under the current test-root contract. Do not use user media, the screenshot's share, repository `.build`, or arbitrary temp paths as test data.

On the planning machine on 2026-09-14 the marked roots were `Z:\RedSalamander.Perf` (repository drive, ReFS, runner default), `C:\RedSalamander.Perf`, and `D:\RedSalamander.Perf`; the machine resource manifest names `D:\RedSalamander.Perf` as the alternate volume, and runs that exercise it require `-AllowAlternateVolumeTestRoot` (BR-6 in `Specs/Plans/Done/FileOperations_BeelineResidualsAndProviderRoutes_2026-09-08.md:226`). Without that switch the managed-directory case, `R4A19_DiscoveryIndependentVolumes`, and the cross-volume host-admission fixture (`Phases05_06.cpp:623`) are skipped or blocked, so every FileOps and Full command below carries it.

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter DiscoveryScope_ -TestRoot X:\RedSalamander.Perf -AllowAlternateVolumeTestRoot
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter FileOpsFamily_DiscoveryScope -TestRoot X:\RedSalamander.Perf -AllowAlternateVolumeTestRoot
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter FileOpsFamily_Phase05_Discovery -TestRoot X:\RedSalamander.Perf -AllowAlternateVolumeTestRoot
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter Phase8_PerItemOrchestration -TestRoot X:\RedSalamander.Perf -AllowAlternateVolumeTestRoot
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter FileOpsFamily_Phase11_BridgeAndConnections -TestRoot X:\RedSalamander.Perf -AllowAlternateVolumeTestRoot
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter FileOpsFamily_R4A19DiscoveryBaseline -TestRoot X:\RedSalamander.Perf -AllowAlternateVolumeTestRoot
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter FileOpsFamily_Fairstream -TestRoot X:\RedSalamander.Perf -AllowAlternateVolumeTestRoot
.\Tools\Run-AllTests.ps1 -Suite Commands -Configuration Debug -CaseFilter cmd_pane_fileops_popup_progress_contracts -TestRoot X:\RedSalamander.Perf
.\Tools\Run-AllTests.ps1 -Suite Commands -Configuration Debug -CaseFilter file_operations_phase0_typed_contract_source_guard -TestRoot X:\RedSalamander.Perf
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Release -CaseFilter FileOpsFamily_DiscoveryScope -SelfTestRepeat 3 -SelfTestShuffleSeed 913 -TestRoot X:\RedSalamander.Perf -AllowAlternateVolumeTestRoot
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -Configuration Debug -TestRoot X:\RedSalamander.Perf -AllowAlternateVolumeTestRoot
git diff --check
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
```

`DiscoveryScope_` step names and the `FileOpsFamily_DiscoveryScope` family are new and must be registered in C0; every other family/case name, runner parameter, and script name above was verified in the current registrations on 2026-09-14. Each candidate command must exit 0 and represent all intended cases; a zero-case run or skipped required deterministic case is not a pass. Release through the runner provides the test-enabled profile. Add `-SkipBuild` only with a valid matching full-solution build receipt. Focused/Affected results cannot replace final Fresh Full.

Preserve the runner's no-activation class unless a real foreground assertion is required. Never kill an independently launched app to free build output. The public spec inventory command writes disposable reports; use its read-only module inventory during planning when only WIP edits are authorized.

## Performance and evidence

- Protected scenarios: single-file time to fixed progress, source bind-to-first-payload latency, managed directory discovery ownership, native merge ownership, and wide-tree traversal/transfer overlap.
- Reuse the terminal `FileOps.Discovery.OpenUs`, `CallbackCount`, `CallbackUs`, `LockWaitUs`, `MaxQueueDepth`, `FirstMutationBeforeClose`, `BytesCompletedWhileOpen`, `MutationsCompletedWhileOpen`, `SkipReleaseUs`, and `Closed` metrics. Add only the missing bounded active-leaf observation in the existing test diagnostics; no per-buffer log stream.
- For a known isolated leaf, prove closure before payload progress. Preserve `FirstMutationBeforeClose=1` for the existing wide-tree overlap scenario; it is not a universal requirement for leaves.
- Compare baseline and candidate on the same machine, profile, fixture sizes, endpoint topology, strategy, and verification setting. Report discovery-open duration separately from transfer throughput. Deliberately held test checkpoints are correctness evidence, not speed measurements.
- Record byte/read/write counts, first-payload timing, callback cost/count, and existing queue/cancel controls. Require no extra whole-file read, no totals pre-scan, bounded discovery memory, and no change to source cleanup/verification outcome. Set any latency/throughput threshold from valid same-machine baseline evidence before evaluating the candidate; do not invent an improvement from the screenshot.
- Keep runtime scratch/evidence generation beneath the owned test root. Archive governed compact evidence through the existing archival flow under `Specs/TestRuns/<MachineHash>/FileOps/` and `Commands/`, preserving results, traces, metrics, case statuses, and baseline/candidate provenance. Follow current archive limits and run `Tools/Test-TestRunArchive.ps1 -RunPath <archive>` and `-Inventory` at closeout.

## Completion checklist

- [ ] C0-C3 implemented; failing-before/passing-after witnesses establish D1-D3 without weakening streaming controls.
- [ ] Known leaf totals close before payload, remain exact, and cover zero bytes, placeholders, retries, mixed roots, verification, and Move cleanup.
- [ ] Nested delete/rename and initial native-directory collision cannot close parent traversal; final closure is idempotent and totals stop growing.
- [ ] Real card/footer/compact/accessibility/taskbar observations agree with task state; ETA retains normal estimator warm-up.
- [ ] Focused Debug, repeated test-enabled Release, and Fresh Full gates pass; required fixture absence is marked blocked with the exact missing prerequisite.
- [ ] Same-machine evidence and archive validation recorded; no performance claim rests only on source inspection.
- [ ] C4 gives every accepted provider/API residual an indexed live successor or an evidence-based rejection.
- [ ] Durable leaf/scope rules and regression obligations merged into the authoritative File Operations, popup, provider, and performance specs. Targets: `FileSystem_FileOperations.md` `## Streaming discovery and scheduling` (`:599`; the leaf rule at `:620-624` is already normative, add the nested-scope ownership rule), `UI_FileOperationsPopup.md` `### Streaming discovery` (`:169`) and `## Progress Rules` (`:398`), `Plugins_VirtualFileSystem.md:1170-1177`, and `Testing_PerformanceValidation.md:745-756`.
- [ ] Only owned files changed; `git diff --check` and spec inventory pass; completed core plan moved to `Specs/Plans/Done/` and I21 removed from the active index.

## Implementation record

Implemented and verified on 2026-09-14 from `63d78f55`. Every anchor in this plan was re-verified against the source before the work began.

### What changed

- `Task::PerItemCallbackCookie` (`RedSalamander/FolderWindow.FileOperationsInternal.h`) gained two discovery-scope fields: `suppressDiscoveryReports`, set once for a cookie a traversal owner hands to its own nested provider calls, and an atomic `deferDiscoveryClosure` for a provisional mutation whose outcome is not yet known. `Task::FileSystemReportDiscoveryProgress` honours the first after header validation and the second at the closure decision only.
- `CrossFileSystemBridge` owns a `nestedProviderCookie` and initialises its `FileSystemOptions` with it, so every provider call the bridge makes inside its own walk carries the suppressed scope. Exact source cleanup, `RelocateChild`'s Native rename, owned-stage creation and removal, readers and metadata all inherit it. The ordinary progress, conflict and result callbacks keep the root cookie, which is passed separately at each call site, so no callback behavior changed. C2's requirement that the adapter outlive the call is met by construction: the cookie is a bridge member.
- `PublishSelectedLeafDiscovery` adds the bound source's exact size to the bridge's cumulative bytes exactly once and reports the root closed, called from `PumpAndPublishFileOnce` immediately after `BindSource` succeeds and before `RouteWriter`. The once-only flag lives on the bridge, not the transaction, because `PumpAndPublishFile` re-runs the whole single attempt on a replace re-prompt. `publishSelectedLeafDiscovery` is threaded from `TraverseAndPublishPath`'s leaf branch and from `CopyLink`'s bound-regular-file placeholder branch, and is false for every walker child.
- A selected root that classifies as a name-surrogate link closes on its already-counted item in `CopyLink`, under both Preserve and Skip, with no target enumeration and no target byte total.
- The Native Move attempt in the per-item executor sets `deferDiscoveryClosure` around the provider call and clears it with a scope guard.
- New observation-only counter `_discoveryGrowthAfterCloseCount`, emitted as `FileOps.Discovery.GrowthAfterClose`. It counts discovery reports that added work after the task's totals were already final. Closure is one-way, so a correct traversal never produces one.
- New self-test pause point `g_fileOpsSelectedLeafDiscoveryPublishedPausePoint`, held immediately after a selected leaf publishes and closes and before any payload work, wired through `State.Private.h`, `State.cpp`, `State.Queue.cpp` and `FileOperationsInternal.h`.
- Two popup debug exports, `DebugResolveFileOperationsWholeTaskProgress` and `DebugBuildFileOperationsTaskHeaderText`, so tests assert the card's own projection instead of recomputing the rules under test. This is the essential observation gap the Scope section allowed.

### Regressions added

`FileOpsFamily_DiscoveryScope` (5 steps, registered in `kFileOpsFamilyDefinitions`, so unfiltered FileOps and Full runs execute it):

| Case | What it proves |
|---|---|
| `DiscoveryScope_LeafCopy` | At a held checkpoint before any payload: the live task reports one file of exactly 3,146,521 bytes with discovery closed, and the real popup snapshot shows a fixed full-file denominator, one item, a determinate meter, no discovery activity, no marquee, and an item header that takes the known-total branch. |
| `DiscoveryScope_LeafMove` | The same through exact Move source cleanup, plus the source is removed only after the transfer. A closed scope is not permission to delete. |
| `DiscoveryScope_ManagedDirectory` | A cross-provider managed directory Move of six children: nested exact source cleanup does not close the walker's root, totals are exact, and growth-after-close is 0. |
| `DiscoveryScope_NativeRenameMerge` | A same-endpoint directory Move refused by an occupied destination continues as a merge walk; the refused attempt does not latch the root closed. |
| `DiscoveryScope_MixedRoots` | A leaf plus a directory in both selection orders: the aggregate closes once on the exact sum. |

### Failing-before and passing-after witnesses

Each fix was disabled in isolation and the family re-run through the governed runner on the same machine and profile.

| Finding | Disabled behavior | Result |
|---|---|---|
| D1 | Leaf never publishes or closes its own scope | `DiscoveryScope_LeafCopy` FAILED: the task finished before the held leaf checkpoint could be observed, because the checkpoint is only reachable once the leaf record is published. |
| D2 | Bridge nested calls report into the selected root | `DiscoveryScope_ManagedDirectory` FAILED: `growthAfterClose=5`, totals grew past the real tree to `files=7/6 bytes=917595/786510` after being presented as final. |
| D3 | Provisional Native rename closes the root | `DiscoveryScope_NativeRenameMerge` FAILED: `growthAfterClose=6` while the merge itself completed correctly, so the defect is presentation truth, not data movement. |
| all | Restored | The whole family passes, `growthAfterClose` is 0 in every case. |

### Notes and deviations

- C2's isolation vehicle is a second `PerItemCallbackCookie` rather than an `IFileSystemOperationControl` adapter. It carries the same operation controls, has no vtable, and cannot outlive the bridge that owns it. `FileSystemShouldAbort` and `FileSystemGetDiscoveryMode` ignore the cookie, so cancellation, deadline and discovery-mode forwarding are unchanged.
- C2's provisional-rename handling defers closure rather than suppressing and republishing counts; see the C2 step above for why the delivered design is simpler and covers the file-rename case as well.
- D3 was confirmed to be premature closure only. The walker publishes its cumulative counts before each nested call, so the cookie's monotonic guard already zeroed nested deltas and no count was ever corrupted.
- The `DiscoveryScope_ManagedDirectory` case uses a cross-provider managed Move, whose source cleanup is the same exact bound delete the plan named, instead of an alternate-volume Local Move. It is deterministic, needs no second volume, and is not a same-volume native rename relabelled as managed. The alternate-volume control remains covered by `R4A19_DiscoveryIndependentVolumes`.
- No totals pre-pass, no extra whole-file read, no new setting, no hardcoded UI string, no ABI layout change, and no change to strategy qualification, verification, link policy or source-cleanup authority.
- Gotcha for the next owner: the governed runner refuses a build receipt if any repository file changes while the build is running. Finish document edits before starting a gate.

### Validation gates

Fresh Full on 2026-09-14 from `63d78f55` plus this working tree: **status passed**, all 20 plan entries promoted, **2,130 passed / 0 failed / 54 declared skips** in 68 minutes, with flaky, regression, isolation-suspect and unclassified-failure classifications all zero. FileOperations rose from 180 to 186 passes for the six new cases. Every skip is environment-gated: absent remote connection profiles, the two 7z cases owned by Compare, two ViewerSpace opt-ins, and two Commands clipboard cases the run could not open the OS clipboard for. The suite-wide invariant holds: `FileOps.Discovery.GrowthAfterClose` is 0 across all 376 recorded task completions. Controls stayed green, including the wide-tree overlap witness, the alternate-volume control and both Beeline native-rename cases. Evidence is archived and validated at `../../TestRuns/4cb089111a23/FileOps/2026-09-14_085731_i21_discovery_scope_fresh_full/README.md`.

### Post-gate closeout step

Moving this plan into `Specs/Plans/Done/` requires exact Done-path admission, which the spec inventory
enforces as a protected-history allowlist. That admission is the only `Tools/` change in this work and
it is deliberately outside the repair scope this plan declared: it adds one path to
`Test-RSProtectedDoneHistory`'s `AllowedAdditionalPaths` in
`Tools/Modules/Tooling/SpecInformationArchitecture.psm1` and the paired positive/negative test the
lifecycle policy requires in `Tools/Tests/SpecInformationArchitecture.Tests.ps1`. No command surface,
inventory classification or governance rule changed, so `Tools/tool-inventory.json` and
`Specs/Testing/Testing_ToolingGovernance.md` need no edit.

Because it lands after the Fresh Full gate, it is separately qualified, following the September 8
closeout precedent for a post-gate metadata change. Its owning test file passes 31/31, the tooling
governance and inventory contracts pass 12/12 and 17/17, and the shared test-harness source contracts
pass 188/188, including the FileOperations Step phase-order drift guard that covers the six new steps
registered here. Closeout gates then read: spec inventory 0 blocking findings, `git diff --check`
clean, and the TestRuns archive contract passing in whole-inventory mode over 4,630 files.

### Residual routing

- D4 is owned by `FileOperations_ProviderDiscoveryParity_2026-09-14.md` (index I22). Route evidence is recorded there: Curl, S3, Google Drive and Dummy all own admitted same-endpoint Copy and Move routes that bypass the bridge and report no discovery, and all four already hold the exact size before mutating. Microsoft Drive is Move-only because it never advertises `copyOperation`. MTP has no size in hand on that route. 7z is rejected on evidence: every mutation entry point returns `ERROR_NOT_SUPPORTED`.
- D5 is owned by `FileOperations_DirectApiDiscoveryContract_2026-09-14.md` (index I23).

## Stop and reconciliation conditions

Pause the affected step and report concrete evidence if the route is no longer as described, source size is not available at the proposed boundary, or an adapter cannot preserve the operation-control lifetime/cancellation contract. Do not improvise an ABI change, reopen a closed task, add a preflight scan, or weaken exact Move cleanup to make tests pass. If a required owned alternate volume, fake backend, or build profile is unavailable, continue independent work and mark only that validation blocked. Unsupported live-provider capability is not authorization to implement it.

Future reviews should check discovery ownership whenever introducing a nested provider mutation, a retry/requalification path, or a new semantic reparse type. Keep active-state tests alongside final-result tests; successful output alone cannot prove truthful progress.
