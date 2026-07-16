# Code Review — `master`, last 2 days (6 commits + working tree)

**Date:** 2026-06-26
**Reviewer:** Claude Code (local max-effort `/code-review ultra`, then a multi-agent deep-verification pass)
**Scope:** `2fc4f817f..HEAD` plus uncommitted working tree — ~57 code files. Docs under `Specs/` excluded.

## Method

1. **First pass** — 10 finder angles (correctness × 5, cleanup × 3, altitude, conventions) + per-candidate verification + gap sweep → 10 findings + 1 watch item.
2. **Deep pass** — a 34-agent workflow that, for every finding: deep-verified it against live source (exact quote, current line, refined severity, concrete fix), then **adversarially tried to refute** the survivors; **re-challenged** the 7 items previously cleared; and ran a fresh **gap sweep** over the new code clusters. Results below supersede the first pass.

### What the deep pass changed
- **F5 strengthened** to the top correctness item — it is now two concrete defects, one of them *deterministic* (read-only attribute lost on **every** overwrite of a read-only file).
- **R6 promoted** from "cleared" to a real **medium** bug (UNC share-root mishandling in the new ancestor-creation walk).
- **F6 and W1 refuted** and dropped (no realistic trigger / self-terminating by SQLite design).
- **8 new findings (N1–N8)** discovered, almost all in the staged-overwrite/promote path that the working-tree commit introduces.
- **F1** confirmed but downgraded to medium (bounded ~5 Hz, self-healing on reset).

## Closeout Update — 2026-06-26

State: **Done**. The review was challenged against the current source, and all meaningful accepted defects were fixed or converted into explicit source/spec contracts.

Implementation closeout:
- **Fixed:** F5, F9, N1, N2, N3, N4, N5, and F3 by replacing the duplicated writer/copy promote paths with one shared staged-promotion helper. The helper clears staged-temp `READONLY` before `ReplaceFileW`, preserves replacement `READONLY` on the final file, strips writer-only hidden/temp attributes, avoids `ReplaceFileW` for reparse replacements, restores destination read-only on failed promotion, restores captured DACL/timestamps after the `MoveFileExW` fallback on best effort, and rolls back staged-copy progress on failed parallel promotion.
- **Fixed:** R6 by treating UNC share roots as ancestor-walk boundaries in resolved destination directory creation.
- **Fixed:** F1 by capping pending content-update and pending subdir-aggregate retry attempts and clearing stale pending work when the cap is reached.
- **Fixed:** F4, F8, F10, N7, and N8 with targeted control invalidation, cached compare details strings, bulk diagnostics trimming, cancelled-result match flushing, and native follow-symlink queue bounding.
- **Fixed/contained:** F7's visible/indexable staged-temp issue by making writer temps hidden/temp and excluding RedSalamander staged temp names from local index hydration. A global crash-orphan deletion sweep was not added because the current implementation has no durable, safe root registry to distinguish RedSalamander-owned stale siblings from arbitrary user files across all folders; per-operation cleanup remains best effort.
- **Rejected as not a bug:** F2 remains the intended cycle-safety behavior. Existing source contracts require followed local directories whose physical identity cannot be probed to be skipped with `FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED`, because traversing without identity would reopen reparse-loop risk.
- **Rejected as not actionable:** N6 was not reproduced as a meaningful leak in the supported staged-copy paths; temp cleanup handles file vs directory attributes and remains best effort for cancellation/failure cleanup.

Authoritative specs updated:
- `Specs/FileSystem/FileSystem_FileOperations.md`
- `Specs/Core/Core_Search.md`
- `Specs/Core/Core_CompareDirectories.md`

Validation evidence:
- Build: `.\build.ps1 -ProjectName RedSalamander` — PASS, 0 warnings/errors. Archived at `Specs/TestRuns/SINON/CodeReview/2026-06-26_MasterLast2DaysCloseout/Build_RedSalamander/`.
- Compare source contract: `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter search_low_hardening_smoke` — PASS, 1 passed. Archived at `Specs/TestRuns/SINON/CodeReview/2026-06-26_MasterLast2DaysCloseout/Compare_search_low_hardening_smoke/`.
- FileOps staged overwrite/floodgate flow: `.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Fairstream` — PASS, 20 passed. Archived at `Specs/TestRuns/SINON/CodeReview/2026-06-26_MasterLast2DaysCloseout/FileOps_Fairstream/`.

Remaining work: none for this review closeout. Broader crash-orphan temp scavenging would require a separate design that records safe sweep roots/ownership before deleting anything across user folders.

## Commits in scope

```
aa89c6013 Add subsystem audit remediation docs
1d1cccc11 Finish FS bridge search index data safety
0e77f10a0 Harden search indexing and document FS followups
ec403afac Fix full-source review bugs
8680621a8 Fix compare and file operation review findings
7613b115c Finalize CompareDirectories remediation closeout
(+ working tree: FileSystem.FileOps.cpp staged-copy-overwrite + self-tests)
```

---

## Priority-ranked findings (most severe first)

| # | Sev | Area | Finding |
|---|-----|------|---------|
| F5 | **medium** | FileSystem write/copy | Staged-overwrite promote drops dest READONLY on every overwrite + ACL/ctime on fallback |
| R6 | **medium** | FileOps resolved-items | UNC share root mishandled in destination ancestor walk → item aborts on network shares |
| F1 | **medium** | CompareDirectories | Content-compare retry busy-loops ~5 Hz when a folder becomes permanently inaccessible |
| N1 | **medium** | FileSystem copy | Read-only source forces every staged overwrite onto the lossy `MoveFileExW` fallback |
| N2 | **medium** | FileSystem copy | Staged symlink overwrite promoted via `ReplaceFileW` (undefined for reparse replacements) |
| F3 | low | FileSystem copy | Spurious `ERROR_PARTIAL_COPY` when source is concurrently resized during a staged copy |
| F2 | low | Search | Follow-symlinks search drops an enumerable subtree when the identity probe fails |
| F4 | low | DxUi | Indeterminate checkbox keeps a stale dash glyph after `VK_LEFT`/`VK_RIGHT` no-op |
| F7 | low | FileSystem | Staged temp siblings are non-hidden, indexable, and orphaned on crash (no sweep) |
| N3 | low | FileSystem copy | Parallel staged-copy failure over-counts global progress by a full file size |
| N4 | low | FileSystem copy | Symlink `ERROR_INVALID_PARAMETER→ERROR_NOT_SUPPORTED` remap masks promote failures |
| N5 | low | FileSystem | `REPLACEFILE_IGNORE_MERGE_ERRORS` silently drops destination ADS / EAs |
| N6 | low | FileSystem copy | `BestEffortDeleteCopiedTemp` leaks a file-symlink temp whose target is a directory |
| N7 | low | Search | Parallel name-only search discards already-found matches on mid-enumeration cancel |
| N8 | low | Search | Plugin `MarkQueuedDirectory` lacks the follow-symlinks queue bound its sibling has |
| F9 | cleanup | FileSystem | Duplicated promoter logic has **already drifted** (divergent unprobeable-dest policy) |
| F8 | cleanup | CompareDirectories | `GetCompareDetailsTextStrings` rebuilds 14 resource strings per item |
| F10 | cleanup | FileOps diagnostics | Pending-flush trims a `std::vector` via repeated front-erase |

**Refuted / dropped:** F6 (WARNING_OVERFLOW mislabel — no realistic trigger, coarse by design), W1 (SQLite re-VACUUM — one-shot, self-terminating; covered by self-test).

---

## Correctness / robustness

### F5 — Staged-overwrite promote silently drops destination metadata  *(medium · refined, survived adversarial refute)*
**Files:** `Plugins/FileSystem/FileSystem.cpp:1884` (`PromoteTempIntoFinalPath`), twin `Plugins/FileSystem/FileSystem.FileOps.cpp:2841` (`PromoteCopiedTempIntoFinalPath`)
```cpp
if (::MoveFileExW(_path.c_str(), _finalPath.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH) == 0)
```
Two distinct defects in both promoters:

- **(A) Deterministic READONLY loss.** When overwriting a read-only destination with the replace-readonly grant, the code clears `READONLY` off the destination before `ReplaceFileW`, then sets `_committed/committed = true` on success. The restore `scope_exit` is gated on `!committed` (`FileSystem.cpp:1873` / `FileOps.cpp:2881`), so it **never re-applies** the bit. **Every** successful overwrite of a read-only file silently leaves it writable.
- **(B) Fallback ACL/creation-time loss.** When `ReplaceFileW` fails (sharing violation from a scanner/second handle, or a filesystem without `ReplaceFile` support), the `MoveFileExW` fallback renames the fresh temp over the destination. The temp carries the parent dir's inherited DACL, a fresh creation time, and no `READONLY` — so the destination's original DACL (possible **ACL widening**), creation time, and attributes that `ReplaceFileW` would have preserved are discarded.

**Fix:** after a successful promote, re-apply the original `READONLY` whenever `clearedReadOnly` is true (regardless of success). On the `MoveFileExW` fallback, capture the destination's security descriptor (`GetNamedSecurityInfoW`) and creation time (`GetFileTime` via a `wil::unique_hfile`) before the rename and re-apply them after (`SetNamedSecurityInfoW` + `SetFileTime`). Apply identically to both promoters (see F9 — fix once by unifying them).

### R6 — UNC share root mishandled in destination ancestor walk  *(medium · promoted from "cleared")*
**File:** `RedSalamander/FolderWindow.FileOperations.State.cpp:8348` (`ensureResolvedDirectoryShell` walk; same flaw in `ensureResolvedDestinationParent` ~8376)
```cpp
for (std::filesystem::path parent = currentPath.parent_path();
     ! parent.empty() && parent != parent.root_path() && parent != currentPath;
```
MSVC `std::filesystem` returns `root_path()` of a UNC path as `\\server\` (the server), **not** the share root `\\server\share`. For any destination under a network share the loop pushes `\\server\share` into `ancestors` and `ensureOneDirectory` runs `GetAttributes`/`CreateDirectory` on the share root. If `GetAttributes` returns a non-missing error (`ERROR_ACCESS_DENIED` / `ERROR_BAD_NETPATH`) the whole item aborts; even on a missing code, `CreateDirectory` cannot create a network share and fails — again aborting. Only the local-share happy path is masked.

**Fix:** treat the UNC share root as a terminating boundary equal to the filesystem root, reusing `TryGetUncShareRoot` (`RedSalamander.cpp:5742`). Add a `parent != shareRoot` guard (path-normalized) to both the ancestor-walk loop and the `ensureResolvedDestinationParent` guard; never `GetAttributes`/`CreateDirectory` a share root.

### F1 — Content-compare retry busy-loops ~5 Hz on a permanently-inaccessible folder  *(medium · refined, survived refute)*
**File:** `RedSalamander/CompareDirectoriesEngine.cpp:2540`
```cpp
static_cast<void>(QueueContentUpdateFolderRetryLocked(folderKey, currentVersion));
return;
```
A folder F that succeeded initially (queuing content-compare jobs into `_pendingContentCompareUpdates[F]`) becomes permanently inaccessible (`FAILED` hr) under a stable version. `ApplyPendingContentCompareUpdatesLocked` hits the cache-miss branch (`:2536`), re-queues a High-priority rescan, and **keeps** the pending entry (the bool is discarded). The rescan recomputes `FAILED`, `IsCacheableDecision` declines to cache (`:3964`), so the miss recurs; `FlushPendingContentCompareUpdatesBudgeted` returns non-empty (`:1115`), re-arming the 200 ms decision-refresh timer. Net: a self-sustaining rescan + UI tick every ~200 ms, ending only on version bump / reset / background-work-disable. `_scanScheduledKeys` dedup caps it to ~one scan per tick (no unbounded queue). The sibling subdir path (`QueueSubdirAggregateRetryLocked`) shares the missing cap.

**Fix:** bound the re-queue with a per-folder attempt counter (e.g. `std::unordered_map<std::wstring,uint32_t> _contentRetryAttempts`, cleared in `ResetCompareStateLocked`). In the cache-miss block, increment; when it exceeds a cap **or** `QueueContentUpdateFolderRetryLocked` returns false, erase `_pendingContentCompareUpdates[folderKey]` (and the counter) to stop the spin. Erase the counter on the success path (after `:2614`). Apply the same cap to the subdir retry.

### N1 — Read-only source forces every staged overwrite onto the lossy fallback  *(medium · new)*
**File:** `Plugins/FileSystem/FileSystem.FileOps.cpp:2889`
```cpp
if (::ReplaceFileW(finalPath.c_str(), tempPath.c_str(), nullptr, REPLACEFILE_IGNORE_MERGE_ERRORS, nullptr, nullptr) == 0)
```
`CopyFileExWithStagedOverwrite` copies the source via `CopyFileExW` (`:2922`) with no attribute-stripping flag, so a read-only **source** carries `FILE_ATTRIBUTE_READONLY` onto the temp sibling. `ReplaceFileW(final, readonlyTemp)` is rejected by Windows for a read-only `lpReplacementFileName` (`ERROR_ACCESS_DENIED`), so the code silently falls back to `MoveFileExW`, taking the metadata-losing path of **F5(B)** with no error reported. `PromoteCopiedTempIntoFinalPath` clears `READONLY` only on the destination (`:2864`), never on the temp.

**Fix:** before promote, clear `FILE_ATTRIBUTE_READONLY` off the staged temp so `ReplaceFileW` can accept it; re-apply `READONLY` to the final path afterward if the source was read-only. Same normalization in the writer twin.

### N2 — Staged symlink overwrite promoted via `ReplaceFileW` (undefined semantics)  *(medium · new)*
**File:** `Plugins/FileSystem/FileSystem.FileOps.cpp:3949`
```cpp
copyHr = CopyFileExWithStagedOverwrite(context, source, destination, COPY_FILE_COPY_SYMLINK, progress, fileBytes, false);
```
`CopyReparsePointInternal`'s symlink-overwrite branch stages a temp that is itself a file symlink, then `PromoteCopiedTempIntoFinalPath` calls `ReplaceFileW(final, symlinkTemp, REPLACEFILE_IGNORE_MERGE_ERRORS)`. `ReplaceFileW` is designed to merge a *regular* replacement's data into the destination while preserving destination metadata; its behavior with a **reparse-point** replacement is undocumented. The destination may end up not being the intended symlink (a copy of the target, or merged metadata) while the call still returns `S_OK`.

**Fix:** don't use `ReplaceFileW` for staged symlink overwrites. Promote via delete-then-`MoveFileExW` (defined for reparse points), or skip staging for symlink overwrites and recreate the link in place after removing the destination.

### F3 — Spurious `ERROR_PARTIAL_COPY` on a concurrently-resized source  *(low · refined, survived refute)*
**File:** `Plugins/FileSystem/FileSystem.FileOps.cpp:2960`
```cpp
if (tempBytes != expectedFileBytes)
{
    return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
}
```
The source size is sampled at `:3780` *before* the copy and passed as `expectedFileBytes` with `verifyTempSize = true` (`:3805`). `CopyFileExW` snapshots the source as-of-copy-time; if the source grew/shrank between the sample and the copy, `tempBytes != expectedFileBytes` → `ERROR_PARTIAL_COPY`, the temp is deleted, and the destination is left unchanged. The same source succeeds via the non-overwrite `CREATE_NEW` path and the reparse path (`verifyTempSize = false`). No data loss; spurious user-facing failure only.

**Fix:** don't compare against the stale pre-copy size. Re-sample the source size immediately after the temp copy and only fail when `tempBytes` is short of the *current* source size, or rely on `CopyFileExW`'s own short-copy failure.

### F2 — Follow-symlinks search drops an enumerable subtree on identity-probe failure  *(low · refined, was medium)*
**File:** `Plugins/FileSystem/FileSystem.Search.cpp:361` (twin `RedSalamander/SearchFallbackEngine.cpp:430`)
```cpp
std::lock_guard lock(runtime.queuedDirectoriesMutex);
runtime.queuedDirectories.erase(visitKey);
```
**Only with `FILESYSTEM_SEARCH_FOLLOW_SYMLINKS` (default off).** For a subdirectory enumerable via the FS backend but whose Win32 identity probe (`CreateFileW FILE_READ_ATTRIBUTES | FILE_FLAG_BACKUP_SEMANTICS`) fails — an ACL granting List/Traverse but denying ReadAttributes, a special reparse directory, or a transient sharing violation — the new code erases the visit key, sets `WARNING_ACCESS_DENIED_SKIPPED`, and returns false, dropping that directory and its **entire subtree** (regression vs the prior `return true`). Surfaced only as a generic `IDS_FIND_WARNING_ACCESS_DENIED`.

**Fix:** a probe failure means loss of cycle-detection for one node, not that the directory is unreadable. On probe failure, keep the visit key, raise the warning, and still traverse (`return true`), or gate the skip only on a genuine `E_ACCESSDENIED`. Apply identically to `SearchFallbackEngine.cpp` so the engines agree.

### F4 — Indeterminate checkbox keeps a stale dash glyph  *(low · confirmed, survived refute)*
**File:** `Common/DxUi/DxUi.Controls.cpp:2619`
```cpp
if (virtualKey == VK_SPACE || virtualKey == VK_LEFT || virtualKey == VK_RIGHT)
{
    _indeterminate = false;
}
```
`Checkbox::OnKeyDown` clears `_indeterminate` without invalidating, then delegates to `Toggle::OnKeyDown`. For `VK_LEFT` when `_checked==false` (or `VK_RIGHT` when `_checked==true`) `Toggle` takes the guarded `if(_checked)`/`if(!_checked)` branch — false — skipping the only `Invalidate(host)` and returning true. `Paint` gates the dash on `_indeterminate`, so the indicator should change but no repaint is queued; the stale dash persists until an unrelated repaint. (`VK_SPACE`/`VK_RETURN` always flip+invalidate; `OnMouseUp`/`OnMnemonic` safe. Refute confirmed arrow keys reach a focused checkbox directly via `DxUi.WindowHost.cpp:2608`.)

**Fix:** in `Checkbox::OnKeyDown`, call `Invalidate(host)` when it actually clears `_indeterminate` (gate on `&& _indeterminate`); invalidations coalesce, so it's harmless on the already-invalidating arms.

### F7 — Staged temp siblings are non-hidden, indexable, and orphaned on crash  *(low · confirmed)*
**File:** `Plugins/FileSystem/FileSystem.FileOps.cpp:2922` (writer twin `FileSystem.cpp:2083`)
The temp sibling (`<final>.rs_copy_tmp_<pid>_<tid>_<counter>` / `.~rs-write-*.tmp`) is created `FILE_ATTRIBUTE_NORMAL`, so `ReadDirectoryChangesW` watchers fire ADDED/MODIFIED and the local search indexer (`LocalSearchIndexCore.cpp` `EnumerateDirectory`/`HydrateDirectorySubtree` apply no name filter and no hidden skip) indexes the partial file as a transient hit. On crash/kill/power-loss mid-copy the `wil::scope_exit` cleanup (`:2947`) never runs and a multi-GB partial sibling is orphaned; **no startup sweep** for `rs_copy_tmp_`/`~rs-write-` exists.

**Fix:** create both temps `FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_TEMPORARY` (after the copy via `SetFileAttributesW`; at `CreateFileW` for the writer); add a hidden/name skip in the indexer's `EnumerateDirectory` next to the `.`/`..` guard; add a best-effort startup sweep of stale siblings reusing `BestEffortDeleteCopiedTemp`.

### N3 — Parallel staged-copy failure over-counts progress by a full file size  *(low · new)*
**File:** `Plugins/FileSystem/FileSystem.FileOps.cpp:3827`
In the parallel copy path, `CopyProgressRoutine` adds byte deltas to `context.parallel->completedBytes` during the temp copy. When the temp copy completes, the full size has already been added. If `verifyTempSize` (`:2960`) or `PromoteCopiedTempIntoFinalPath` (`:2972`) then fails, `returnFailure` only re-reports progress and does **not** subtract the already-counted `fileBytes`. The reported total exceeds reality by one full file size per failed staged item (worse than the pre-staging direct-write path, which counted only partial bytes).

**Fix:** on a staged-overwrite promote/verify failure, roll back the bytes already added for this item from `parallel->completedBytes` before returning, mirroring the success-path delta accounting (`:3836-3840`).

### N4 — Symlink error remap masks promote failures  *(low · new)*
**File:** `Plugins/FileSystem/FileSystem.FileOps.cpp:3970`
```cpp
if (copyHr == HRESULT_FROM_WIN32(ERROR_INVALID_PARAMETER))
{
    copyHr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}
```
`CopyFileExWithStagedOverwrite` can return `ERROR_INVALID_PARAMETER` from the **promote** step (e.g. `ReplaceFileW`/`MoveFileExW` on an unusual volume), not from the symlink copy. The blanket remap rewrites it to `ERROR_NOT_SUPPORTED`, telling the user "symlinks not supported" even though the symlink staged copy succeeded and only the atomic promotion failed.

**Fix:** scope the remap to the actual non-overwrite symlink `CopyFileExW` call (`:3951`); have the staged path return its promote errors unremapped.

### N5 — `REPLACEFILE_IGNORE_MERGE_ERRORS` silently drops ADS / EAs  *(low · new)*
**File:** `Plugins/FileSystem/FileSystem.FileOps.cpp:2889` (writer twin `FileSystem.cpp:1881`)
Overwriting a destination carrying alternate data streams or extended attributes (`Zone.Identifier`, classification EAs) with a source lacking them: `ReplaceFileW` with `REPLACEFILE_IGNORE_MERGE_ERRORS` ignores failures merging the destination's secondary streams/EAs, dropping them while still returning success — so the caller believes metadata was preserved.

**Fix:** consider dropping `REPLACEFILE_IGNORE_MERGE_ERRORS` so merge failures surface, or at minimum `Debug::Warning` when a merge error is reported so silent ADS/EA loss is observable. Decide deliberately and document.

### N6 — `BestEffortDeleteCopiedTemp` leaks a file-symlink temp pointing at a directory  *(low · new)*
**File:** `Plugins/FileSystem/FileSystem.FileOps.cpp:2832`
```cpp
if (attributes != INVALID_FILE_ATTRIBUTES && (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0u)
{
    static_cast<void>(::RemoveDirectoryW(tempPath.c_str()));
    return;
}
```
A staged symlink copy can create a file-type symlink temp pointing at a directory. On cleanup, `GetFileAttributesW` **follows** the link and returns `FILE_ATTRIBUTE_DIRECTORY`, so the code calls `RemoveDirectoryW` on a file-type reparse point — which refuses — and never attempts `DeleteFileW`, leaking the `.rs_copy_tmp_` sibling.

**Fix:** read attributes without following the link (the link itself reports `FILE_ATTRIBUTE_REPARSE_POINT`) and treat a reparse point as a file (`DeleteFileW`), or fall back to `DeleteFileW` after `RemoveDirectoryW` fails.

### N7 — Parallel name-only search drops accumulated matches on mid-enumeration cancel  *(low · new)*
**File:** `Plugins/FileSystem/FileSystem.Search.cpp:568`
```cpp
return SUCCEEDED(result.status) && ! result.matches.empty();
```
`AccumulateParallelNameOnlyEntry` sets `result.status = kCancelledHr` on cancellation (`:537/:545`). `QueueParallelDirectoryResultChunk` gates emission on `HasQueuedParallelDirectoryResultWork`, which requires `SUCCEEDED(result.status)` — so name matches already accumulated for that directory before the cancel point are silently dropped. The serial path emits as it goes, so scalar vs parallel walks return different visible result sets on cancellation.

**Fix:** on cancellation, flush the matches already accumulated before honoring the cancel, or decouple the `SUCCEEDED(status)` gate from match emission.

### N8 — Plugin `MarkQueuedDirectory` lacks the follow-symlinks queue bound its sibling has  *(low · new)*
**File:** `Plugins/FileSystem/FileSystem.Search.cpp:339`
With `FOLLOW_SYMLINKS`, the plugin engine accumulates one visit key + one identity per distinct directory in unbounded `unordered_set`s. The sibling `SearchFallbackEngine.cpp::MarkQueuedDirectory` guards the same growth with `kMaxQueuedFollowSymlinkDirectories` (1,000,000) and raises `WARNING_OVERFLOW`. On a filesystem exposing an enormous number of distinct directories (deep symlink/junction fan-out across mounts) the plugin engine can drive the worker toward memory exhaustion with no overflow warning — inconsistent safety limits for the same scenario.

**Fix:** add the same `kMaxQueuedFollowSymlinkDirectories` cap + `WARNING_OVERFLOW` signal to the plugin's `MarkQueuedDirectory`.

---

## Cleanup / simplification / architecture

### F9 — Duplicated staged-overwrite promoter logic has already drifted  *(cleanup · confirmed)*
**Files:** `Plugins/FileSystem/FileSystem.cpp:1840` (`PromoteTempIntoFinalPath`) vs `Plugins/FileSystem/FileSystem.FileOps.cpp:2841` (`PromoteCopiedTempIntoFinalPath`)
The two ~50-line promoters are near-duplicates that have **diverged**: on an unprobeable final path (`GetFileAttributesW` fails with `ERROR_ACCESS_DENIED`/`ERROR_SHARING_VIOLATION` rather than NOT_FOUND), the copy path returns the error early (`FileOps.cpp:2852`) but the writer path sets `destinationExists=true` with `attributes` still `INVALID`, skips the readonly block, and lands on a bare `MoveFileExW(MOVEFILE_REPLACE_EXISTING)` — blindly replacing a destination it couldn't even stat. The temp-sibling name generators and the 32-attempt `CREATE_NEW`/`FAIL_IF_EXISTS` retry loops are likewise duplicated.

**Fix:** hoist a single `FileSystemInternal::PromoteStagedTempIntoFinalPath` (both TUs already use `namespace FileSystemInternal` and share `FileSystem.Internal.h`); move the safer FileOps body (early-return on unprobeable destination) in, have the writer wrapper propagate the result into `_committed`, and unify the temp-name generators. This is the right home for the F5 / N1 / N5 fixes — fix the staging primitive **once**.

### F8 — `GetCompareDetailsTextStrings` rebuilds 14 resource strings per item  *(cleanup · confirmed)*
**File:** `RedSalamander/CompareDirectoriesWindow.cpp:342` (call site `:2553`)
`FolderView::RefreshDetailsText` loops over every visible item; `BuildDetailsTextForCompareItem` calls `GetCompareDetailsTextStrings()` per item, and the call site binds `const auto` (a copy). Each call does 14 `LoadStringW` + 14 `std::wstring` allocations producing identical immutable labels → 14·N lookups + 14·N allocations per refresh sweep. The original (commit `74fb8b023`) cached these in a function-local `static` returned by const reference.

**Fix:** restore a function-local `static const CompareDetailsTextStrings` returned by const reference; change the call site to `const auto&`. Keep owning `std::wstring` storage.

### F10 — Diagnostics pending-flush trims a `std::vector` via repeated front-erase  *(cleanup · refined)*
**File:** `RedSalamander/FolderWindow.FileOperations.State.Diagnostics.Part.cpp:906`
```cpp
_diagnosticsPendingFlush.erase(_diagnosticsPendingFlush.begin());
```
`_diagnosticsPendingFlush` is a `std::vector<TaskDiagnosticEntry>` (header `:644`) while the adjacent `_diagnosticsInMemory` is a `std::deque` trimmed via `pop_front`. `erase(begin())` left-shifts every surviving entry (each holding 3 `std::wstring`s). Steady state is one O(n) erase per insertion; the O(n²) shape only appears once when the cap shrinks via a settings reload. `n` is bounded (~256), never crashes (`maxPendingFlush` can't be 0).

**Fix:** one bulk `erase(begin(), begin()+overflow)`, or switch the container to `std::deque` to mirror `_diagnosticsInMemory`.

---

## Refuted / dropped this pass

- **F6 — `WARNING_OVERFLOW` mislabel** *(refuted)*: `MaterializeEntryPaths` failing for a degenerate empty path is essentially unreachable in practice, and the warning is coarse by design. No realistic trigger; not worth a flag split.
- **W1 — SQLite re-VACUUM loop** *(refuted)*: `PRAGMA auto_vacuum=INCREMENTAL` + `VACUUM` persists the mode to the DB header; `PopulateStoreInfo` reads it back on a fresh connection so `ShouldQueueAutomaticMaintenance` returns false next idle window. One-shot and self-terminating, proven by `search_service_sqlite_legacy_auto_vacuum_queues_idle_maintenance`.

## Re-challenged previously-cleared items — outcomes

| Item | Result |
|------|--------|
| R1 `CompareDirectoriesWindow::Create()` menu reorder | **still cleared** — correct on every path; self-destruct case still fires the release guard, avoiding double `DestroyMenu`. |
| R2 `FactoryImpl.h` empty-span returns | **still cleared** — all three templates zero/null out-params before any early return. |
| R3 `DxUi.Accessibility` `scoped_lock` | **still cleared** — `GetAccessibilityTargetMutex` is a `std::recursive_mutex`, the only mutex in the subsystem; no callback held under the lock. |
| R4 SQLite maintenance WAL-checkpoint coverage | **still cleared** — `shouldFinalCheckpoint` set true by every WAL-dirtying branch; all-flags-false early-returns `S_FALSE` before any write. |
| R5 FileOps perf-counter atomics | **still cleared** — `PerfStats` compiler-enforced non-copyable; global sweep found only `.load(acquire)`/`.fetch_add(relaxed)` + a by-reference bind. |
| **R6** UNC ancestor walk | **PROMOTED → medium defect** (now finding R6 above). |
| R7 `OnLeaveScopePrompt` repost-in-`scope_exit` | **still cleared** — reposts require fresh navigation; deferred-delete invariant keeps members live; guarded by `&& _hWnd`; flag cleared first so it never sticks. |

---

## Suggested order of attack

1. **F5 + N1 + F9 together** — unify the two promoters into one `FileSystemInternal::PromoteStagedTempIntoFinalPath` and fix attribute/ACL/readonly handling once (also resolves N5's logging gap and the F9 divergence). This is the highest-value, highest-density fix.
2. **R6** — small, deterministic break on network shares; add the UNC share-root boundary to both guards.
3. **F1** — the only finding that can spin indefinitely; add the retry cap (and mirror to the subdir path).
4. **N2 / N6 / N4** — symlink staging correctness (undefined promote, cleanup leak, error masking).
5. **F4 + F8** — tiny, clearly-correct, user-visible.
6. **F3, F7, N3, N7, N8, F10, F2** — robustness/cleanup as time permits.
