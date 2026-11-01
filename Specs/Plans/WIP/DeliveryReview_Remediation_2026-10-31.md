# I57 Delivery review remediation 2026-10-31

Status: ACTIVE. Implementation: not started. Renumbered from the drifted I56 label to I57 on 2026-10-31: origin/main already assigns I56 to Unified CI hosted acceptance, and plan ranks are never reused.
Priority: high. This plan makes the dirty remediation on `ej/review-remediation-20261031` safe to merge.
It is a merge gate for this worktree. It is not a claim that `origin/main` is broken today.

Planned at `6b6f5646be297fbc232d137f1df6a00c5cb74461` plus the uncommitted remediation already in this worktree.
Review-end `origin/main` was `e361efec1e27ce0a980334309c900c6f285161f5`. Product C++ between that commit and
HEAD is unchanged. Current main writes schema-1 Move breadcrumbs and shows them. The hide exists only in
this dirty worktree, which writes schema 2 (`kSchemaVersion = 2` in `RedSalamander/FileOperationMoveBreadcrumb.h`).

Drift command, before any edit. Re-read every cited line in the dirty bytes. If a line moved, cite the current line.

```powershell
git diff HEAD -- RedSalamander/FileOperationMoveBreadcrumb.h RedSalamander/FileOperationMoveBreadcrumb.cpp RedSalamander/FolderWindow.FileOperations.State.Runtime.cpp RedSalamander/FolderWindow.FileOperationsInternal.h RedSalamander/FolderWindow.FileOperations.cpp RedSalamander/FolderWindow.FileOperations.State.cpp RedSalamander/FolderWindow.FileOperations.State.Private.h RedSalamander/FileOperationDurableStore.h Specs/FileSystem/FileSystem_FileOperations.md Plugins/FileSystem/FileSystem.cpp Plugins/FileSystem/FileSystem.vcxproj
```

Do not revert, restage, or discard the existing dirty remediation to make the tree match HEAD.
Do not start from receipt `b9666fb6` as if it were this tree.

## Ownership

This plan owns three outcomes: schema-1 breadcrumb projection (W2), the interrupted-notice task-id lock (W1),
and the clean-receipt process gate (W3). I36 and I50 stay the owners of their other residuals. Do not copy
this work into those plans as a second implementation, and do not close them when this plan finishes.

Authoritative contract to update with the repair: `Specs/FileSystem/FileSystem_FileOperations.md`.
The implementation change, when it happens, records its own `Changes/` fragment under
`Specs/Core/Core_Changes.md`. This plan-only document does not add one.

## W2 — High. Schema-1 Move breadcrumbs stay invisible and fill the store cap

### Current behavior

`FileOperationMoveBreadcrumb.h` line 16 sets `kSchemaVersion = 2`. `Breadcrumb::Create`
(`FileOperationMoveBreadcrumb.cpp` lines 483-485) forces that version, `GetCurrentProcessId()`, and the
process creation FILETIME. `Serialize` (lines 220-221) refuses a record without a nonzero PID and start time.

`Parse` (lines 400-421) still accepts schema 1 and schema 2. The schema-2 branch reads `writerProcessId`,
`writerStartedFileTime`, and `singleItemName`. A schema-1 record leaves `writerProcessId` and
`writerStartedFileTime` at the `Record` defaults of 0 (header lines 60-61) and leaves `singleItemName` empty.
HEAD's schema-1 writer never stored those three fields. Its JSON fields are `kind`, `schemaVersion`, `taskId`,
`createdFileTime`, `phase`, `sourcePane`, `destinationPane`, `nativeItems`, `managedItems`, `copyOnlyItems`,
`sourceRootCount`, `sourceRootSamplesTruncated`, `sourceRootSamples`, and `destination`.

`ObserveWriter` (lines 75-90) returns `Unverified` when the PID or the start time is 0 (lines 79-80).
`LoadInterrupted` (lines 564-573) counts `Active` and `Unverified` and continues, without appending the
record and without calling `Acknowledge`. `Acknowledge` (lines 657-658) returns `ERROR_BUSY` for anything
that is not `Exited`. A schema-1 file is therefore never shown and never removed.

Those files can already be on disk under `%LocalAppData%\RedSalamander\State\FileOperations\MoveBreadcrumbs`
(`ResolveRoot`, lines 45-46). Main has written schema 1 since the 2026-10-26 snapshot and shows non-terminal
records. After this dirty change they stay invisible for the life of the store.

They still count toward the 4,096-child cap. `kMaximumStoreEntries` is 4,096 (header line 19).
`EnumerateDirectRegularFiles` (`FileOperationDurableStore.h` lines 351-361) counts every direct child other
than `.` and `..`, then returns `ERROR_FILE_TOO_LARGE` and clears the list. `LoadInterrupted` (lines 530-540)
then withholds every notice. The separate visible batch of 24 (`kMaxCompletedTaskSummaries` in
`FolderWindow.FileOperations.State.Private.h` line 196) is a different cap.

The dirty spec documents this bug. The sentence is now line 2977 of
`Specs/FileSystem/FileSystem_FileOperations.md`, in the paragraph that begins at line 2973:

> Unverifiable writers and legacy schema-1 records are retained without projection or acknowledgement and reported in load diagnostics; absence of writer proof never establishes interruption.

The review saw that sentence near 2970-2972. Lines 2970-2972 are now the terminal persist/remove sentences
and stay as they are. The 2026-10-31 writer section is now lines 3418-3431 (the review's 3413-3432). It says
creation always stamps a PID and that acknowledgement refuses active and unverifiable records. Read together,
those sentences leave every schema-1 file unshown and unacknowledged. That text does not justify the hide.
Lines 3409-3412, the ReadOnlyDeleteBlocked apply-to-all paragraph, stay unchanged.

### Repair

On load, project a non-terminal schema-1 record when no other `RedSalamander.exe` can still be its writer.
Acknowledge a terminal schema-1 record under that same condition, using the existing terminal-removal path,
so a completed schema-1 file does not occupy the 4,096 cap forever. Keep the file untouched, and do not
project or acknowledge it, while another instance might still be the writer.

"No other writer" means all of the following:

- The scan looks for other processes whose image file name is `RedSalamander.exe`. The search service is not
  a breadcrumb writer. This process is not a schema-1 writer: `Create` writes schema 2 only. Exclude the
  current PID.
- No such peer is alive, or every such peer has a creation FILETIME strictly after the record's
  `createdFileTime`. A process that did not exist when the record was written cannot be its writer.
  Use the same `OpenProcess` / `GetProcessTimes` observation schema 2 already uses.
- If process enumeration fails, or any peer's creation time cannot be read, the record stays hidden.
  That peer might still be the writer.

Do not weaken schema-2 behavior. `ObserveWriter` for schema 2 stays as written: a live matching PID is
`Active` (skipped, including a terminal record still being finalized); a recycled PID whose creation time
differs is `Exited` (lines 87-88); access or identity-query failure, other than `ERROR_INVALID_PARAMETER`
for a dead PID, and a wait that is neither exit nor timeout, stay `Unverified`. `Parse` still rejects a
schema-2 record with a zero PID (lines 415-418). Do not treat every zero PID as projectable. Only a
successfully parsed schema-1 record takes the legacy path.

Do not rewrite a schema-1 file into schema 2, and do not invent a PID or a `singleItemName`. Schema 1 has
no display name. A one-item schema-1 notice keeps an empty subject. The fixture below uses more than one
admitted item so the existing empty-name rule for a multi-item notice applies.

A schema-1 record withheld because a peer might be the writer must not increment `unverifiedWriterRecords`
and must not make the partial-load warning claim that writer proof failed. Add a separate load count for
that withhold. Schema-2 unverified records stay on the existing counter, stay out of the projected vector,
and stay unacknowledged. `Acknowledge` of a schema-2 record that is not `Exited` still returns `ERROR_BUSY`.

Once a schema-1 record is safe to project, dismissal uses the existing `Acknowledge` path and removes that
exact direct child. A later `LoadInterruptedMoveBreadcrumbs` can then fill the visible batch from whatever
safe record remains.

### Spec

Replace the schema-1 half of the line 2977 sentence so the contract matches the repair. Schema-2 unverified
writers stay retained without projection or acknowledgement, and absence of schema-2 writer proof still
never establishes interruption. Schema-1 records are projected, or removed when terminal, only when no
other `RedSalamander.exe` might be the writer; otherwise they stay on disk, unprojected and unacknowledged,
and still count toward the 4,096-child cap.

Update lines 3418-3431 so "creation always stamps a PID" and "acknowledgement refuses unverifiable records"
are the schema-2 rules, with the schema-1 rule beside them. Leave lines 2970-2972 and 3409-3412 as they are.
Leave line 3431's clear-completed rule as it is: clear completed preserves every notice.

### Acceptance check that fails on these bytes

Extend `RunFileOpsInterruptedMoveWriterChecks` in
`RedSalamander/SelfTest/FileOperations/FileOperations.Scope.Tests.ReviewRemediation.cpp`. `Breadcrumb::Create`
cannot write the fixture: it stamps schema 2. Write a schema-1 `*.move.json` into the self-test store with
HEAD's field set, a non-terminal phase, and more than one admitted item.

- With no other `RedSalamander.exe`, `LoadInterrupted` returns that record in the interrupted vector, does
  not count it as unverified, and a `FileOperationState` projects one notice. `Acknowledge` removes the file.
  On these bytes the record is skipped at lines 570-573 and `Acknowledge` returns `ERROR_BUSY`.
- Start a second `RedSalamander.exe` with `CREATE_SUSPENDED` before writing the fixture, so its creation
  time is at or before `createdFileTime`, and leave it alive for the check. The record stays out of the
  interrupted vector, the file remains, and `Acknowledge` does not remove it. Close the process afterwards.
  A peer started after `createdFileTime` is not a possible writer, so that record is projected. The check
  needs both peers: one that might still be the writer, and one that cannot be.
- Schema 2 stays as specified. The live-writer assertions at lines 192-194 still require an empty load,
  `activeWriterRecords == 1`, and `Acknowledge` returning `ERROR_BUSY`. The recycled-PID assertions at
  lines 195-197 still project the record. A schema-2 record whose `OpenProcess` fails with an error other
  than `ERROR_INVALID_PARAMETER` still increments `unverifiedWriterRecords`, stays out of the vector, and
  is refused by `Acknowledge`. Do not use a zero PID as that fixture.

## W1 — Medium. Interrupted-notice task ids race the queue-order counter

### Current behavior

`LoadInterruptedMoveBreadcrumbs` (`FolderWindow.FileOperations.State.Runtime.cpp` line 664) assigns
`summary.taskId = _nextTaskId++` with no lock. `_nextTaskId` is a plain `uint64_t`, initialized to 1
(`FolderWindow.FileOperationsInternal.h` line 2938). The function locks `_mutex` only to count existing
notices (lines 603-606) and to push the summary (lines 701-704).

These allocators increment under `_mutex`:

- admission, lines 1448-1451 of the same file
- batch rename, `FolderWindow.FileOperations.cpp` lines 3103-3106
- informational cards, lines 3516-3536 of the runtime file
- `AllocateLateQueueOrderKey`, lines 2856-2859, whose comment at internal header lines 2570-2572 says it
  consumes an id from the admission counter

`DismissCompletedTask` (lines 3194-3230) holds `_mutex` while it erases the summary, drops the lock,
acknowledges, and then calls `LoadInterruptedMoveBreadcrumbs`. A file-operation task thread can be inside
`Task::PrepareForExecution` (line 11759), which calls `RunSameHostOverlapAdvisory` (line 11831). That
function calls `AllocateLateQueueOrderKey` at line 11541 when the task queues after a later peer.

The unlocked increment and the locked increment are a C++ data race on `_nextTaskId`. The cross-thread
consumer is a fresh queue-order key. This race does not dismiss the wrong history row and does not collapse
two queued tasks. Do not describe it, and do not test it, as either of those failures.

### Repair

Increment under the same `_mutex` as the other allocators, then push the summary under that same hold.
Keep the assignment `summary.taskId = _nextTaskId++` so the compacted source guard in
`Commands.Tests.PluginConfig.cpp` lines 5565-5578 still matches. Do not hold `_mutex` across the breadcrumb
directory read or the notice string formatting. Do not replace the mutex with a lone atomic increment:
the other allocators already publish the counter under `_mutex`.

There is no reader for the counter. A test-only getter under `ENABLE_TESTS`, next to
`DebugLoadInterruptedMoveBreadcrumbsForSelfTest` (internal header lines 2715-2717), is enough. It is not
a product behavior change.

### Acceptance check that fails on these bytes

Run `AllocateLateQueueOrderKey` on one thread and `LoadInterruptedMoveBreadcrumbs` on another, using enough
projectable breadcrumbs that the unlocked line 664 overlaps the locked line 2859. Every returned queue-order
key and every projected `summary.taskId` is unique. After both threads join, the counter equals one plus the
number of ids actually handed out. On these bytes the increment at line 664 is outside `_mutex`, so a lost
update fails that equality. Do not assert a wrong dismissed row or a collapsed queue.

## W3 — High. Process gate

Do not merge until a clean x64 Debug rebuild of the bytes being merged succeeds with zero warnings and
zero errors, and `FileOps_ReviewRemediationContracts` then runs on that build receipt. The receipt's source
identity has to be those bytes, including the W1 and W2 repairs.

Keep the failed run `20261031T110058Z-63848-18d88350187d422dbd06b888af75f49e`. I50 records it at
`CodeReview_MainMerges_2026-10-28.md` lines 130-134: ViewerPE export-close failures, PluginContract exit
`0xC0000005`, and tooling failures. That run is an older receipt. It is not evidence that the access
violation still reproduces. Do not delete it, and do not rerun it as a substitute for the receipt above.

Drop the idea of rerunning `FileOps_ReviewRemediationContracts` on receipt
`b9666fb669301ecf77e7a84428f8e9deb17839e0f39ff245cfb8b9677a0d22a7` and calling that coverage. I50 lines
123-129 say that receipt, from `.build/logs/msbuild-20261031_115410_085-pid74156-87f0c8ab.log`, precedes
the later ViewerSpace, provider, breadcrumb, and lifecycle candidates and cannot qualify them. The focused
pass on that receipt (`20261031T105821Z-74180-398861ecf36946c7ac243bbc03fdef24`) stays historical.

Two later logs are not receipts for these bytes:

- `.build/logs/msbuild-20261031_130217_670-pid28696-92239977.log` failed because `volumeRoot` and
  `fallbackDriveRoot` were undeclared, and `emittedRoot` was an uninitialized reference, at the then-current
  `Plugins/FileSystem/FileSystem.cpp` line 1546. Current line 1546 is the unknown-storage fallback and does
  not name those identifiers. `volumeRoot` and `fallbackDriveRoot` still exist as parameters of
  `ResolveStorageProbeDevicePath` at lines 1409-1425. The broken call site is gone.
- `.build/logs/msbuild-20261031_131853_200-pid77540-65de65f8.log` failed with `LNK2019` for
  `FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem` and `CompareFileSystemPathsByComponents`.
  `Plugins/FileSystem/FileSystem.vcxproj` line 124 now compiles `Common/FileSystemPathIdentity.cpp`.

A later log that links `RedSalamander.exe` is still not this gate unless it is a clean governed receipt of
the bytes being merged and the file-operations remediation contracts have run on that receipt.

## Follow-on, not blocking W2

DxUi pin `4cdf1ce52c7785446698e346b8aec058ed6a9283` (`Dependencies/DxUi.lock.json`) lacks the menu
hover-timer keyboard cancel that exists only on DxUi `origin/main` commit `762983a0`. That commit is
R28-45 in DxUi `Specs/Plans/WIP/CodeReview_2026-10-28.md`. A later pin bump points at that DxUi
delivery-review plan. It is not part of this remediation and does not block W2.

No RedSalamander product call site uses the DxUi Source view. The editor bugs, including Source-view Enter
(DxUi `Specs/Plans/WIP/RichTextAudit_2026-10-21/findings.md`, editor-03), stay in the DxUi plans until a
pin bump. Make File List `sourceMode` is a different setting.

## Preserve

These behaviors stay as they are in the dirty remediation:

- Synchronous host payloads fail closed.
- Conflict defaults for delete and same-host overlap stay as they are.
- Leftover cleanup does not preselect in-use rows.
- Schema-2 writer observation distinguishes a recycled PID (`ObserveWriter` lines 87-88, and the
  self-test at `ReviewRemediation.cpp` lines 195-197).
- Clear completed keeps notices (spec line 3431, and the history-eviction check in the same self-test).

## Closed — do not reopen

The 2026-10-31 delivery review closed these. They are not work items in this plan.

- Wrong-row dismissal and collapsed queue keys as the impact of W1. The impact is the data race above.
- ViewerSpace `ScanMain` use-after-free. The scan lambda AddRefs; `AbandonScanWorkers` bumps the
  generation; `PostUpdate` drops a stale generation; `OnDestroy` does not free the object while `ScanMain`
  is inside it.
- MTP direct create replacing an existing occupant. The upload refuses an existing leaf before create.
  A partial object if Cancel does not delete is unproven.
- Publication workflow printing the winget token. It passes the token on stdin.
- Delete apply-to-all granting a generic access-denied delete. It still requires a fresh
  `ReadOnlyDeleteBlocked` receipt. Leave spec lines 3409-3412 unchanged.
- Same-host overlap Enter starting the dangerous action, and permanent-delete Enter deleting. Those
  match the popup spec.

## Verification

Use `Test-Changes.ps1 -Explain` and the affected file-operations scope while editing. The merge gate is
the clean x64 Debug receipt in W3, then `FileOps_ReviewRemediationContracts` on that receipt. The W2 and
W1 checks above have to fail on the pre-repair bytes and pass on that receipt. Preserve the failed run
named in W3. This plan does not authorize a Full run, a focus-taking run, or a DxUi pin bump.

## Done

Move this plan to `Specs/Plans/Done/` only after W1 and W2 are in the tree, the spec sentences match the
repair, the W3 receipt and contract run are recorded, and the exact new Done path is admitted. I36 and I50
stay open for whatever they still own.

## STOP

Stop if a schema-1 record would be shown or acknowledged while another `RedSalamander.exe` might still be
its writer, if schema-2 live or unverified results change, if the task-id repair is written up as a
wrong-row or collapsed-queue fix, if the only offered evidence is run
`20261031T110058Z-63848-18d88350187d422dbd06b888af75f49e`, receipt `b9666fb6`, or either non-receipt log in
W3, or if the change starts a DxUi pin bump. Keep the failed run. Do not revert the dirty remediation.
