# File Operations: provider discovery parity on direct routes

## Status and ownership

- Status: COMPLETE; implemented on 2026-09-15 and closed on a passing Fresh Full on 2026-09-15 (machine `4cb089111a23`). "Closeout status and how to finish" records how the gate was finished; then see the Implementation record.
- Index owner: I22.
- Priority: P2. No data-safety exposure; the defect is a long-lived indeterminate card on routes that already know their totals.
- Effort: M per provider, delivered as one session for the four that remained.
- Change risk: medium; it changes what each provider publishes to the host during a live transfer, not what it moves.
- Planned at: `63d78f55`, 2026-09-14. Opened by I21 (`FileOperations_DiscoveryScopeAndLeafProgress_2026-09-13.md`) as the routed owner of its D4 residual. Curl and Dummy delivered on 2026-09-14 in `bff0ad9d`; S3, Google Drive, Microsoft Drive and MTP on 2026-09-15.
- Dependencies: I21 owned the host bridge and the host-side discovery contract; I23 the Local provider's direct API. Neither was re-opened here. Shared File Operations files were coordinated with I18 and I3.
- Authoritative contracts: `Specs/Plugins/Plugins_VirtualFileSystem.md`, `Specs/FileSystem/FileSystem_FileOperations.md`, `Specs/FileSystem/FileSystem_Mtp.md`, `Specs/UI/UI_FileOperationsPopup.md`, `Specs/Testing/Testing_PerformanceValidation.md`.

This plan is self-contained and non-normative. The audit behind it was source-based and every claim below about a
live provider comes from a fake backend, not a live account or device.

## Closeout status and how to finish

Everything a fresh session on any machine needs is in this file and in git; nothing lives only in a
session or a machine.

### Where the work stands (2026-09-15, `master`)

- Slices landed, in order: `90c1b47a8` (Local reader synchronous EOF), `b1a0460ac` (recording fixtures; Curl
  and Dummy bulk roots and leaf closure), `9df48b9a7` (S3), `f03e6c948` (Google Drive), `00842e94e`
  (Microsoft Drive), `2ac262c8b` (MTP), `f5b900abe` (source guard and spec merge), `5777d5e26` (this record
  moved to Done), then the commit carrying this section, which re-bases two FileOps controls on the provider
  contract (see Corrections).
- Focused witness on the delivered tree: `FileOpsFamily_DiscoveryScope` 15/15 on `LT-PF5VDAGE`
  (`7d3a1247382a`), run `20260915T113812Z-43164-d3a526c144c84c37b7e4a30739a6e96b`; the raw run directory
  was reclaimed by a later run's stale-run cleanup.
- Local gates on the final tree: `git diff --check`, `Tools/Get-SpecInventory.ps1 -FailOnFindings`,
  `Tools/Tests/SpecInformationArchitecture.Tests.ps1`, `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`.

### What the Fresh Full on LT-PF5VDAGE showed

Run `20260915T115831Z-23456-f9334d03856143af8b164dbbfad45dae` on tree `5777d5e26` (build receipt
`aaf65dc611e67d1f310bc8b23c754a5b601735515c7b93d50f9b4edaaded48b2`, workspace snapshot
`a928e27f9b8e5f8434ae6d33d7b0314f0cb00dd6d41181c973b66db26c927021`, plan digest
`58336f513976e44ce58d3aabd7df254fc5020e40c27108305ad67ae98bf6d344`): status failed; 1,298 counted, 1,168
passed, 19 failed, 111 skipped, in 139 min 32 s. Its validation checkpoint (run-state, plan, per-entry
results and terminal results) survives at `D:\RedSalamander.Perf\evidence\runs\<run id>` on that machine;
the raw run directory was reclaimed. The nineteen failures, classified:

1. Two were this plan's. `Phase5_DiscoverySkipContinues` ("Traversal closed before the post-Skip
   just-in-time state could be observed") and the Dummy control of `R4A19_DiscoveryProviderControls` (its
   open-discovery facts) encoded the direct-route behavior this plan removes; both are re-based (Corrections)
   and pass in isolation on the final tree:
   Phase05 family 10/10 (run `20260915T152601Z-22064-25f629c32c114c2ea0e64f9e4e8060d0`), R4A19 family 5/5
   (`20260915T153302Z-22064-de17338514dc4cf9a402fe4833cef0e6`, alternate `C:` root authorized).
2. `Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker`, then the first case of every later FileOps
   family (`R4A19_DiscoveryProviderControls`, `DiscoveryScope_LeafCopy`, `R3_1_IdentityLessReplaceDummy`,
   `R4T_DeepTreeCopyCompletesIteratively`, `R4A02_DeleteOverCopySourceWarnsAndQueues`,
   `Phase6_PopupSmokeResizeAndPause`, `Phase7_CrossPaneVisibleRefreshDummy`,
   `Phase8_DefaultBandwidthLimitFromSettings`, `Phase9_ConflictPrompt_TypeMismatchNoOverwrite`; 12 FileOps
   failures and 81 "not reached" skips in all) and `PluginContractTests`: the session held no
   `SeCreateSymbolicLinkPrivilege` (not elevated, Developer Mode off), so `FileSystem.dll`'s debug
   self-tests fail 105/1 at the object-binding symlink fixture (`RunDebugObjectBindingSelfTest`, "relative
   file-symlink fixture should be created"; the count matches that check exactly, and the plugin's
   `Debug::Error` text is invisible without an ETW listener). Riptide runs those self-tests in-process, and
   in that run every later family's first case then timed out. On the final tree, in isolation from a
   non-elevated shell:
   R3OwnedPublication 7/7 (`20260915T155324Z-22064-eb75fa07d25343bda942d0a5996c84b8`), R4Traversal 5/5
   (`20260915T160059Z-22064-d6a7c61eee864ca4975a6673fe43ba11`); R4Overlap and Phase06 to Phase09 were not
   re-run before this record was written.
3. `C0_CurlCopyTraversalTruth` 141/142, the same in isolation
   (`20260915T154015Z-22064-d980f617a00d41d4a9a4fafe586cf936`): the Wide 4,096-file serial `CopyItem`
   variant hits the corpus's 180 s deadline at 3,813 files (`ERROR_TIMEOUT`) on this laptop's loopback, while
   the 4-way variant completes in 96 s and the whole export takes 311 s against the 186 s the other machine
   records. This plan did not touch the directory walk that route uses.
4. `FileSystemS3.dll`'s internal cancel self-tests in `PluginContractTests` (a streaming request returns 15 s
   after Cancel; "the recursive Delete must list the prefix on the fixture"): reproduced with the parent
   commit's plugin sources, so pre-existing on this machine's loopback.
5. Commands: one access violation in `cmd_pane_embedded_terminal_command_state_truth`; the command palette
   destroys its window on deactivation and `DebugSetSearch` then writes to the dead search field (crash
   report `RedSalamander-20260915-150000-p11724`); something took foreground activation during the run. A
   latent bug outside this plan; see Residuals.
6. ToolsPesterTests 4: `HistoricalTerminalProfile.Tests.ps1` sealed hashes, because this checkout carries
   CRLF working copies of `eol=lf` Terminal test files (`git ls-files --eol Tools/Tests` shows `i/lf w/crlf`),
   and the two reparse tests, the same missing privilege.

### How to finish

1. Get a green Fresh Full on a clean checkout of the final tree. Prefer machine `4cb089111a23`, where I23's
   Full passed on 2026-09-14 (its record names the `Z:`/`D:` roots). On `LT-PF5VDAGE` it takes an **elevated**
   PowerShell, re-materialized Terminal test files (`git ls-files --eol Tools/Tests | findstr w/crlf`, delete
   those working copies, `git checkout -- Tools/Tests`), and the loopback-bound S3 and C0 failures above are
   expected to stay, which is why the other machine is the safer bet. Command:
   `.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -Configuration Debug -AllowAlternateVolumeTestRoot`.
   Do not touch the tree until it finishes; do not launch it from a session that lacks the privilege.
2. Archive before running anything else. From `<test root>\runs\<run id>\artifacts\selftest\last_run\`
   copy `results.json`, `run-all-tests-results.json`, `run-state.json`, `validation-plan.json`,
   `workspace-final.json`, `trace.txt`, `fileops\results.json`, `fileops\trace.txt`, and the
   `FileOps.Discovery.*`, `FileOps.SelfTest.DiscoveryScope.*` and `FileOps.Mtp.Discovery.SourceLookupUs` rows of
   `perf\perf_metrics.jsonl` as `perf/discovery-scope-metrics.jsonl`, into
   `Specs/TestRuns/<machine hash>/FileOps/<yyyy-MM-dd_HHmmss>_i22_provider_discovery_parity_fresh_full/`.
   Write its README on the model of
   `Specs/TestRuns/4cb089111a23/FileOps/2026-09-14_181429_i23_direct_api_discovery_fresh_full/README.md`:
   source commit, run id, command, receipt, snapshot, plan digest, verdict with the per-suite table, the six
   `DiscoveryScope_*DirectApi` witnesses and the controls named under "Recording fixtures (D0)",
   `FileOps.Discovery.GrowthAfterClose` 0 in every sample, the `FileOps.Mtp.Discovery.SourceLookupUs` summary
   (48 to 82 us cache-served on the fake device here), and the declared skips. Validate with
   `Tools\Test-TestRunArchive.ps1 -RunPath <dir>` and `-Inventory`; `git add -f` the directory.
3. In this record replace `EVIDENCE_PLACEHOLDER` with the archive path, set the Status line to COMPLETE with
   the date, tick the last completion-checklist item, and correct the closing date in the dependency rule of
   `Specs/Plans/WIP/README.md`.
4. `git diff --check`, `.\Tools\Get-SpecInventory.ps1 -FailOnFindings`, then one closeout commit
   ("Close I22 on a passing Fresh Full"), as I23 did in `a6db5bde5`.

### How it was finished (2026-09-15, `4cb089111a23`)

Steps 1 to 4 were carried out as written, on the recommended machine, from a clean tree at `41f288e99`
with the symlink privilege confirmed by probe before launch. The run passed on the first attempt there;
its verdict, per-suite table, witnesses, invariant and declared skips are in the archive named in the
completion checklist. Every one of the nineteen LT-PF5VDAGE failures classified above as a machine
condition was absent on this machine, which is the strongest confirmation of that classification
available without a second laptop.

## Problem and required outcome

I21 gave the host bridge an exact leaf discovery record and repaired nested-scope ownership. That fixes every route the bridge executes. It does not reach the routes where a provider is itself the traversal owner, because no provider except the built-in Local one reported discovery at all.

The host bypasses the bridge whenever the operation stays inside one endpoint and needs no managed cleanup or verification. A same-endpoint Copy with verification off calls the provider's `CopyItem` directly, and a same-endpoint Move qualified as Native calls `MoveItem` directly and can never route through the bridge, because Native is excluded from verification planning. On those routes the task's discovery scope stayed open until the per-item terminal guard closed it at the end of the whole item, so the card showed `Discovering` and a transferred-bytes-shaped total for the entire transfer.

Required outcome, now met: every provider that owns an admitted direct route publishes cumulative discovery for that governed call and closes it exactly once, with an exact byte total wherever the provider already holds the size before it mutates. Unknown-size semantics are preserved; no provider reports an unknown file as a known zero-byte file.

## Implementation record

### What every provider now does

The shared owner is `Common::FileOperations::DiscoveryScope` (`Common/FileSystemDiscoveryScope.h`, catalogued in
`Specs/Core/Core_SharedHelpers.md`). Every provider constructs it from the call's own options, never from the
thread-local ambient carrier each of Curl, S3, Google Drive and Microsoft Drive installs, hoists it to the public
entry point so a helper shared by singular and bulk routes only adds, and closes a singular call the moment its
totals are final. The durable rule and the per-provider table now live in `Specs/Plugins/Plugins_VirtualFileSystem.md`
("Every shipped provider that mutates on a route the host calls directly ...").

| Provider | Delivered | Where |
|---|---|---|
| Curl | 2026-09-14, plus 2026-09-15: the bulk entry points (`CopyItems`, `MoveItems`, `DeleteItems`, `RenameItems`) now note each selected root on their pre-pass. The 2026-09-14 slice reported roots only from the singular entry points, so a bulk call closed on totals short by its roots; the D0 fixture found it on its first run. | `FileSystemCurl.Internal.h` (`FileOperationProgress::discovery`), `FileSystemCurl.CopyMove.cpp` |
| Dummy | 2026-09-14, plus 2026-09-15: the bulk walks (`CopyItems`, `MoveItems`) report one discovery per node they visit, `DeleteItems` the subtree each root removes, `RenameItems` the node it renames. Same D0 finding as Curl. | `FileSystemDummy.cpp` |
| S3 | 2026-09-15. `ExecuteCopyOrMove` hands the complete plan to a `reportDiscovery` observer right after `BuildTransferPlan`, before any object work; a singular call adds and closes there. Bulk calls publish every root from the existing `EstimateTransferBytes` pre-pass, which already built each plan, and close before the first transfer; the per-root helper is then not asked again. `DeleteResolvedPath` reports the object, or the first listing of the prefix it clears. `ClassifyPlannedObjects` treats a key ending in `/` as a directory and a listed prefix as one directory even with no marker. | `FileSystemS3.Directory.cpp` |
| Google Drive | 2026-09-15. `DriveTransferContext` carries the scope pointer; `DriveTransferPath` takes the scope and whether it may close (singular yes, batch no). A leaf and a native folder move close before the transfer; `DriveTransferTree` reports the root and every child at the point it is popped from its frame, which covers the collision-skip branch. `DriveDeletePath` and the identity-pinned delete report the item alone. `GoogleItem::sizeKnown` records whether the API reported a size, so a native document is an unknown-size file, not an empty one. `RenameItem`/`RenameItems` ride the same path. | `FileSystemGoogleDrive.cpp` |
| Microsoft Drive | 2026-09-15. The plan's P5 premise was wrong at the current code: `MoveOrRenameItem` already fetches the source item's metadata (`:4556`), which carries `isFolder` and `sizeBytes`. So Native Move reports an exact leaf, or one directory, with no added request; a merge reports each child it visits or skips. `DeleteItem`'s body became the shared `DeleteSingleItem` so `DeleteItems` keeps one scope for all its roots instead of letting each item close the call's scope. | `FileSystemMicrosoftDrive.cpp` |
| MTP | 2026-09-15. Decision recorded below. `IMtpBackend::GetCachedItemSummary` answers kind and last-listed size from the path cache without invalidating it; `FileSystemMtp::ReportSourceDiscovery` runs it as its own read-only command, carries the result out through shared state, and reports before the mutating command. `CopyItem`/`MoveItem` share `CopyOrMoveSingleItem`, `DeleteItem`/`DeleteItems` share `DeleteSingleItem`, so `CopyOrMoveItems` and `DeleteItems` keep one scope. | `FileSystemMtp.Internal.h`, `FileSystemMtp.Device.cpp`, `FileSystemMtp.FakeBackend.cpp`, `FileSystemMtp.h`, `FileSystemMtp.Core.cpp` |
| 7z | Rejected on evidence, unchanged: every mutation entry point is `ERROR_NOT_SUPPORTED`. | |

### The MTP decision

The plan asked for the size question to be decided on evidence. The evidence, from `FileSystemMtp.Device.cpp`:

- The WPD backend's own `CopyItem`/`MoveItem` resolve the source through `ResolvePathCached` and already hold
  `item.attributes` and `item.sizeBytes` at that point; the plugin core never saw them because the backend
  interface returned only an HRESULT.
- `GetAttributes` is served from the path cache. `GetFileSize` deliberately invalidates the cache first, because it
  owns the live committed size, so it always costs a device round trip. It is the wrong tool for discovery.
- `RunBackendCommand` detaches a command the watchdog abandons and returns to the caller, so a command may never
  hold a pointer into the caller's frame; results leave through shared state, as `tempPuidMissing` already does.

Decision: one cache-served lookup, `GetCachedItemSummary`, run as its own read-only command before the mutation.
A warm cache, the normal case after the listing the user is looking at, costs no device round trip; a cold cache
pays the same path resolution the mutation would have paid next, moved earlier. A file reports its last-listed
size, a directory one directory, because the device relocates or copies it as one object. The cost is recorded
per lookup in `FileOps.Mtp.Discovery.SourceLookupUs` (detail `served=cache|device|none`), so a live device can
be measured without re-deriving this; on the fake backend the lookup is the dispatch of one command (see the
evidence archive for the figure). This is more than the item-based record the plan admitted as a first step,
and it needed no size probe.

### Corrections that changed the design while implementing

- Curl and Dummy were not fully delivered on 2026-09-14: singular routes reported, bulk routes closed short.
  The first D0 fixture run named it; both are fixed above. The plan's statement that Curl "avoids [per-item
  closure] by construction" was right about closure and wrong about roots.
- Microsoft Drive does hold the size before it mutates (see P5 above), so it reports exact bytes.
- A pre-existing Local provider defect blocked every FileOps case on this machine: the overlapped reader returned
  a synchronous `ERROR_HANDLE_EOF` from `ReadFile` as a failure, while the pending route already mapped it to
  the zero-byte EOF the spec promises. The bridge always issues one read at end-of-file, and on this Dev Drive
  volume that read completes synchronously, so every Local-sourced bridge copy failed with EOF before its stage
  was created. `Plugins/FileSystem/FileSystem.cpp` now maps the synchronous case the same way; committed on its
  own.
- A pre-existing S3 defect, outside this plan's scope and left as a routed task: a same-bucket prefix transfer
  fails with `ERROR_REVISION_MISMATCH` whenever the prefix holds the marker object `CreateDirectory` writes,
  because the marker's key is re-resolved as a plugin path and the bucket is dropped from its HEAD (the fake's
  request log shows `HEAD /discovery-scope/src/tree/ -> 404`). The S3 fixture therefore seeds its tree as an
  implicit prefix, which is how the plugin's own S3 self-tests seed prefixes.

- The Fresh Full found two FileOps controls whose expectations encoded the direct-route behavior this plan
  removes. `Phase5_DiscoverySkipContinues` pressed Skip on a Dummy tree Copy and required the traversal to
  still be open afterwards; `R4A19_DiscoveryProviderControls` required the delayed Dummy Copy to complete
  bytes and mutations while discovery was open. Both were satisfied only because the host's per-item terminal
  guard used to close the Dummy scope at the end of the item (and the 2026-09-14 Dummy closed at destruction,
  after its throttled progress). A Dummy that closes on its exact subtree before the first byte leaves Skip
  nothing to switch and completes nothing while open. The Skip case now drives the Local walker, whose
  bounded discovery-ahead queue and 64 B/s transfer keep the traversal producing (the `discovery-a`/`b`
  trees the mode-transition cases already seed), and the Dummy control requires the closed exact record
  (512 KiB, 64 files, 9 directories) with zero open-discovery bytes/mutations and zero growth after closure;
  the fake-MTP control keeps the overlap facts because the host bridge traverses its route.
  `Specs/Testing/Testing_SelfTests.md` and `Specs/Testing/Testing_PerformanceValidation.md` say so.

### Recording fixtures (D0) and the host witness

`RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.DiscoveryProviders.cpp` adds one case
per provider to `kFileOpsFamilyDiscoveryScope`: `DiscoveryScope_DummyDirectApi`, `_CurlDirectApi`, `_S3DirectApi`,
`_GDriveDirectApi`, `_GraphDirectApi`, `_MtpDirectApi`. Each starts the provider's fake backend through its
self-test export, seeds three leaves of odd sizes and one flat tree through the provider's own writer, drives the
entry points under a recording `IFileSystemOperationControl`, and asserts per call that the cumulative stream never
decreases, that exactly one report closes it, that the closure is last, and that it closed on the exact expected
totals. Every call runs even after an earlier defect, so one run names them all. Each case then runs one
same-endpoint leaf transfer through the host on the admitted direct route and requires the completed task to
carry a closed exact record with no mutation completed while discovery was open. The shared pieces
(`DirectApiDiscoveryRecorder`, `DescribeDiscoveryScopeDefect`, `RunDirectApiDiscoveryCall`, `FakeBackendProvider`,
`SeedDiscoveryFixtureSelection`, `DescribeHostLeafDiscoveryDefect`) live in the main FileOps self-test file, and
the Local case now uses them too. The fake's request log is appended to a failure so the provider's actual
backend traffic explains itself.

Failing-before witnesses, from the family runs on 2026-09-15 (see the evidence archive):

- Curl and Dummy, delivered code: bulk calls "closed on wrong totals" short by their roots.
- S3: `CopyItem(leaf) reported no discovery at all. | MoveItems(leaves) reported no discovery at all. |
  DeleteItems(leaves) reported no discovery at all.`
- Google Drive, Microsoft Drive and MTP: every call "reported no discovery at all" at the delivered code.

### D4 contract guard and spec merge

`Tools/Tests/TestHarnessSourceContracts.Tests.ps1`, "gives every provider mutation entry point on a direct route
one discovery scope": every `Plugins/FileSystem*` directory must be described or exempt, every Copy, Move and
Delete entry point of each described provider must carry the scope or delegate to a helper that does, and no
provider may call `FileSystemReportDiscoveryProgress` itself; only the shared helper does. The Local provider
keeps its own aggregator by design and is guarded by the sibling Local test. Rename entry points are outside the
guard: the host's rename route carries no operation control.

`Specs/Plugins/Plugins_VirtualFileSystem.md` carries the durable obligation and the per-provider table;
`Specs/FileSystem/FileSystem_Mtp.md` the backend summary lookup. The built-in capability-v2 profile matrix has no
discovery column and no capability advertisement changed, so the plan's conditional matrix note resolved to no
change.

## Prioritized findings

Route evidence taken at `63d78f55`. `useCrossFileSystemBridge` is `(_destinationFileSystem != nullptr || hasManagedMovePlans || hasVerificationPlans) && (COPY || MOVE)` in `RedSalamander/FolderWindow.FileOperations.State.cpp`; `_destinationFileSystem` is null exactly when the endpoint tuple matches (`RedSalamander/FolderWindow.FileOperations.cpp:3298`); Native is excluded from verification planning (`State.cpp:17463`). Direct provider calls: `CopyItem` at `State.cpp:18913`, `MoveItem` at `State.cpp:18950`.

| ID | Provider and route | Size known before mutation | Disposition |
|---|---|---|---|
| P1 | Curl leaf Copy and Native rename. | Yes. | **Done 2026-09-14**, bulk roots 2026-09-15. |
| P2 | S3 object Copy and Native rename. Capability at `FileSystemS3.Core.cpp:634` for standard S3 profiles. | Yes, and for the whole prefix: `BuildTransferPlan` produces the plan before any object work. | **Done 2026-09-15.** Exact closed total up front. |
| P3 | Google Drive server Copy and Native Move. Capability at `FileSystemGoogleDrive.cpp:2918` when the profile is writable. | Yes: `GoogleItem::sizeBytes` resolved before `DriveTransferFile`; unknown for native documents, now recorded as such. | **Done 2026-09-15.** |
| P4 | Dummy delayed Copy and Native Move. | Yes on the node; an ungenerated subtree is one directory. | **Done 2026-09-14**, bulk walks 2026-09-15. |
| P5 | Microsoft Drive Native Move. `CopyItem` returns `ERROR_NOT_SUPPORTED` and the profile never advertises `copyOperation`. | Yes, contrary to the plan: the move's own metadata GET carries `isFolder` and `sizeBytes`. | **Done 2026-09-15.** Exact leaf, one directory for a folder, no added request. No direct Copy was added. |
| P6 | MTP Copy, and the CopyOnly Move that falls back to `CopyItem` because MTP advertises no `nativeMoveOperation`. | From the backend's path cache, yes; from the device, only at a round trip. | **Done 2026-09-15.** Cache-served summary by its own command; decision and evidence above. |
| P7 | 7z. Every mutation entry point returns `ERROR_NOT_SUPPORTED` and the profile advertises neither copy nor move. | n/a | Rejected on evidence. Archive export runs through the host bridge and is covered by bridge semantics. |

## Scope

- One provider per step, each with its own fake-backend fixture and its own focused verification.
- `Specs/Plugins/Plugins_VirtualFileSystem.md` for any durable provider obligation, and this plan.

Capability advertisement, strategy qualification, transfer and verification algorithms, link policy and public
ABI layout were not changed. No unsupported provider feature was implemented. No totals pre-pass was added to a
provider that did not already hold the total; S3's bulk pre-pass and MTP's path cache already did.

## Execution checklist

### D0 — Fixture observation before any producer change

- [x] A recording operation control that records every discovery report, shared by every provider case.
- [x] For each provider row, the admitted direct route driven and the event stream recorded; the baseline was zero reports for the four remaining providers and short bulk totals for the two delivered ones.
- [x] The live task's discovery state asserted, not only the final result: `discoveryMutationsCompletedWhileOpen` must be zero on the host leaf route.

### D1 — Curl and Dummy

- [x] Publish the leaf's exact size and close the governed call's scope before the first byte moves.
- [x] Extend Curl's recursive walkers to report cumulative totals as they enumerate, closing once at true completion.
- [x] Preserve the sizeless-listing path: when `sizeKnown` is false the record stays item-based with no byte total.
- [x] Recording fixtures for both, per D0, which found and fixed the bulk-root gap.

### D2 — S3 and Google Drive

- [x] S3 reports the plan's exact total once, before object work begins; scope hoisted to `CopyItem`, `MoveItem`, `CopyItems`, `MoveItems`, `DeleteItem` and `DeleteItems`; the helper only adds. Prefix semantics and the ancestor-collision guards unchanged. S3 never needs the unknown-size form.
- [x] Google Drive reports the leaf exactly and lets `DriveTransferTree` own recursive totals through a scope pointer on the const context; each child reported when popped from its frame; the native folder move reports one directory.
- [x] Both proved against their fake backends. Cancellation while discovery is open is covered by the existing `R0f*Containment` cancel witnesses, which the scope does not alter: `FileSystemShouldAbort` stays live and closure on cancellation is the destructor's job.

### D3 — Microsoft Drive and MTP

- [x] Microsoft Drive Native Move publishes one closed record per item, exact for a leaf; the scope is constructed after the header validation the helper delegates.
- [x] MTP decided on evidence; the decision and its measurement are recorded above.

### D4 — Contract and closeout

- [x] Provider-contract test that fails a provider which mutates on an admitted direct route without ever reporting discovery.
- [x] Durable obligation merged into `Specs/Plugins/Plugins_VirtualFileSystem.md` and `Specs/FileSystem/FileSystem_Mtp.md`.

Verified after each step: the popup shows a fixed denominator and a determinate meter for a known leaf on that provider's direct route (the host witness in every case), and `FileOps.Discovery.GrowthAfterClose` stays 0.

## Machine and gate facts, corrected for this closeout

The plan was written on a machine with a `Z:` ReFS root and a `D:` NTFS alternate. On `LT-PF5VDAGE` there is no
`Z:`; the marked root is `D:\RedSalamander.Perf` on a ReFS Dev Drive, which is also the runner's default because
the repository is on `D:`, and the only alternate fixed volume is `C:` (NTFS), which `-AllowAlternateVolumeTestRoot`
lets the FileOps suite mark as `C:\RedSalamander.Perf`. So the commands below carry no `-TestRoot`.

Gate traps, all paid for during I23's closeout on 2026-09-14, plus one paid for here:

- The working tree is shared with other sessions and a Fresh Full validates the **working tree**, not your commits.
- A commit landing on master mid-run forces the run's exit code to 1 (`Tools/Run-AllTests.ps1`, `SOURCE MUTATION`).
- **Any** edit to the tree while the build phase runs, a spec or a Pester file included, makes the build refuse to
  publish a receipt ("Repository source changed while the build was running") and the run fails before testing.
- Interrupting a run leaves `.build\artifact-operations\*.contaminated.json` and the next run is refused until `build.ps1 -Rebuild` clears it.
- Archive the evidence **before** running anything else; a later run's stale-run cleanup reclaims `runs\<id>`.
- `Specs/TestRuns/*` is gitignored; an evidence archive must be `git add -f`'d.
- Moving this plan to `Specs/Plans/Done/` needs its path in `AllowedAdditionalPaths` in `Tools/Modules/Tooling/SpecInformationArchitecture.psm1` and a paired test in `Tools/Tests/SpecInformationArchitecture.Tests.ps1`.
- A session without `SeCreateSymbolicLinkPrivilege` cannot pass the FileOps suite or PluginContractTests
  (see "Closeout status and how to finish"); check `whoami /priv` or `New-Item -ItemType SymbolicLink` before
  spending two hours on a Full. A killed governed run leaves a stale `.build\artifact-operations\*.lock`
  whose owner is gone; the next build or test marks the profile contaminated and only a full-solution
  `build.ps1 -Rebuild` repairs it.
- A governed FileOps run sweeps the run directories of dead runner processes but keeps those of a live one,
  so several `-CaseFilter` runs launched from one PowerShell process keep each other's `runs\<id>` until that
  process exits.

## Commands and verification gates

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter FileOpsFamily_DiscoveryScope -AllowAlternateVolumeTestRoot
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter FileOpsFamily_Phase05_Discovery -AllowAlternateVolumeTestRoot
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter FileOpsFamily_R4A19DiscoveryBaseline -AllowAlternateVolumeTestRoot
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -Configuration Debug -AllowAlternateVolumeTestRoot
git diff --check
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
```

## Completion checklist

- [x] Every row above is implemented, or rejected with evidence recorded here.
- [x] Each implemented provider has a failing-before and passing-after witness on its fake backend.
- [x] No provider reports an unknown size as a known zero-byte total.
- [x] Cancellation, deadline and discovery-mode behavior are unchanged on every touched route.
- [x] Durable provider obligations merged into `Specs/Plugins/Plugins_VirtualFileSystem.md`.
- [x] Fresh Full passes; `git diff --check` and the spec inventory pass; the completed plan moves to `Specs/Plans/Done/` and I22 leaves the active index. The move and the local gates were done on 2026-09-15; the Fresh Full on `LT-PF5VDAGE` failed for the machine reasons recorded under "Closeout status and how to finish", and the gate was then finished the same day on `4cb089111a23`, the machine that record recommended. Evidence: `../../TestRuns/4cb089111a23/FileOps/2026-09-15_183715_i22_provider_discovery_parity_fresh_full/README.md`.

## Residuals routed elsewhere

- S3 prefix transfer over a folder marker object (`ERROR_REVISION_MISMATCH`): routed as a standalone task, not
  owned here; transfer algorithms were out of this plan's scope.
- Rename entry points carry no operation control on the host's rename route and are outside the D4 guard; Curl,
  Dummy and Google Drive report on them anyway because their shared helpers made it free.
- Command palette use-after-free (Commands suite): `CommandPaletteWindow` resets its window on
  `WM_ACTIVATE`/`WA_INACTIVE` and detaches the DxUi host on `WM_NCDESTROY` without clearing `_searchField`,
  `_grid`, `_model` or `_root` (`RedSalamander/CommandPaletteWindow.cpp`), so `DebugSetSearch` writes to a dead
  `TextField` when anything takes activation while `cmd_pane_embedded_terminal_command_state_truth` runs. The
  palette should null those pointers when its window goes and the debug hooks should return false; the case
  should report the lost activation instead of crashing.
- `FileSystem.dll` debug self-tests need `SeCreateSymbolicLinkPrivilege` for the object-binding symlink
  fixtures and fail 105/1 without saying so, because the local `check` lambdas only call `Debug::Error`, which
  is silent without an ETW listener; `Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker` and
  PluginContractTests then fail in every non-elevated session. Either probe the privilege and declare a
  skip/blocker, or mirror those failures to stderr the way `Common::DebugSelfTest::Check` does.
