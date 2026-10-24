# Code Review — `master` delta re-verify (fixes for 2026-06-26 findings + new commits)

**Date:** 2026-06-27
**Reviewer:** Claude Code (local max-effort `/code-review ultra`, multi-agent verify+hunt workflow, 50 agents)
**Delta under review:** `aa89c6013..HEAD` (`a37fdde7f`) — the 5 commits that landed after the prior review:

```
a37fdde7f Fix cross-FS move source cleanup safety
acfdc1a4e Fix staged promotion review findings
0f9860ffc Update master code review findings
773fa0ad3 Stabilize FileOps bridge aggregate selftest
452b99c84 Finish FS deep audit remediation slice
```

**Method:** verified each prior-finding fix is correct/complete (Phase 1), hunted the new code for fresh bugs with 6 finder angles weighted to the new cross-FS move source-cleanup surface (Phase 2), adversarially verified all candidates (Phase 3), gap-swept (Phase 4), synthesized (Phase 5). The critical finding (N1) was additionally re-confirmed by hand against live source.

## Bottom line

**Closeout update (2026-06-27):** all real findings below were implemented from high to low priority. DR-12 was latent/plausible rather than independently reproduced in the review, but the cleanup path was hardened anyway because the fallback is low-risk. No meaningful implementation item remains open from this review.

**Post-closeout validation update (2026-06-28):** follow-on validation found remaining meaningful failures in the test and cleanup surfaces, not new unresolved review findings. Commands GUI selftests now use owner-scoped DxUi popup lookup/dismissal, cross-thread-safe menu debug getters, message-pumped popup-driver joins, focused NavigationView edit re-resolution after refresh, and warning/restore guards around the remaining real cursor probes. Compare selftests now treat only `ERROR_INVALID_WINDOW_HANDLE` from host-owned connection UI as a documented pre-window precondition skip. FileOps cross-FS directory Move cleanup now still enumerates changed non-reparse source directories and deletes copied children that remain proven unchanged, while preserving the changed directory itself and failing closed for changed reparse directories.

Validation completed:
- `git diff --check` passed (line-ending normalization warnings only).
- `Invoke-Pester -Path Tools\Tests\TestHarnessSourceContracts.Tests.ps1` passed: 32 passed / 0 failed.
- `.\build.ps1 -ProjectName RedSalamander` passed in Debug x64 with 0 warnings / 0 errors. Build log: `.build\logs\msbuild-20260627_125311_348.log`.
- Focused runtime selftests passed: `Phase12_ReparsePointPolicy`, `Phase11_CrossFileSystemBridge`, `Floodgate_LocalWriterOverwriteIsStaged`, `Floodgate_LocalCopyOverwriteIsStaged`, `search_low_hardening_smoke`, and `search_local_plugin_parallel_cancel_fanin`.
- 2026-06-28 revalidation: `.\build.ps1 -ProjectName RedSalamander` passed with 0 warnings / 0 errors (`.build\logs\msbuild-20260628_150339_986.log`); full Commands passed `762 / 0 / 2` (`Specs/TestRuns/7d3a1247382a/Commands/2026-06-28_140401/commands_results.json`); full CompareDirectories passed `162 / 0 / 29` (`Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-28_143051/compare_results.json`); full FileOps passed `101 / 0 / 20` (`Specs/TestRuns/7d3a1247382a/FileOps/2026-06-28_152421/fileops_results.json`).

---

## Fix-verification status (prior findings)

| Fix cluster | Prior findings | Status | Note |
|---|---|---|---|
| FV-retrycap | F1 | ✅ CORRECT | Retry caps (`kMaxPendingUpdateRetryAttempts=3`) + symmetric pending/retry-map clearing fully correct; no leak/desync. |
| FV-small | F4, F8, F10 | ✅ CORRECT | Indeterminate-clear invalidate, magic-static compare strings, bulk vector trim — all correct and safe. |
| FV-promote | F5, F9, N1(old), N5(old) | ✅ FIXED | Post-commit attributes are warning-only after promotion, `ReplaceFileW` again uses `REPLACEFILE_IGNORE_MERGE_ERRORS`, and MoveFileEx fallback restores only destination security/creation time while preserving replacement last-write/access times. |
| FV-symlink-progress | N2, N3, N4, F3 (old) | ✅ FIXED | Symlink copy `ERROR_INVALID_PARAMETER` remaps to `ERROR_NOT_SUPPORTED` in overwrite and non-overwrite paths; staged overwrite progress rollback now handles sequential and parallel copies. |
| FV-tempvis | F7, N6 (old) | ✅ FIXED | Staged-temp cleanup has reparse fallback deletion, and the local indexer now recognizes generated temp-name shapes instead of raw marker substrings. |
| FV-unc | R6 | ✅ FIXED | UNC share-root matching now normalizes slash style before boundary comparison. |
| FV-search | N7 (old) | ✅ FIXED | Parallel result flushing still emits accumulated matches but no longer clears a fatal directory status. |

## Remediation status

| Finding | Status | Resolution |
|---|---|---|
| DR-1 | ✅ Fixed | Cross-FS MOVE cleanup now detects source directory reparse points, does not enumerate through them, deletes only the link object after verification, and Phase12 covers `followTargets` with an out-of-tree target sentinel. |
| DR-2 | ✅ Fixed | Promoted files return `S_OK` after commit; final attribute failures are logged with `Debug::Warning` only. |
| DR-3 | ✅ Fixed | `QueueParallelDirectoryResultChunk` only resets `result.status` when the queued status succeeded, preserving fatal terminal errors. |
| DR-4 | ✅ Fixed | Move cleanup avoids full content re-download for high-latency/cloud/virtual endpoints and falls back to full compare when storage classification cannot be obtained. |
| DR-5 | ✅ Fixed | `ReadersHaveEqualContent` now compares stream bytes across independent short-read boundaries instead of requiring equal per-call read sizes. |
| DR-6 | ✅ Fixed | Retained `moveBridge` has its callback cookie re-bound before cleanup in both per-item loops. |
| DR-7 | ✅ Fixed | Cleanup probes use bounded retry for transient HRESULTs and return retryable cleanup failures instead of silently converting them to permanent partials. |
| DR-8 | ✅ Fixed | `StagedPromotionOptions::ignoreReplaceMergeErrors` defaults on and feeds `REPLACEFILE_IGNORE_MERGE_ERRORS`. |
| DR-9 | ✅ Fixed | MoveFileEx fallback restores destination DACL/creation time only, not stale destination last-write/access times. |
| DR-10 | ✅ Fixed | UNC share-root guard now compares normalized separators. |
| DR-11 | ✅ Fixed | Symlink copy invalid-parameter remap no longer depends on overwrite mode. |
| DR-12 | ✅ Hardened | Temp cleanup now falls back between `RemoveDirectoryW` and `DeleteFileW` for reparse ambiguity. |
| DR-13 | ✅ Fixed | Search/index temp filter validates generated suffix patterns for bridge/copy/writer temps. |
| DR-14 | ✅ Fixed | Sequential staged-copy rollback restores `context.completedBytes` to the per-item base. |
| DR-15 | ✅ Fixed | Cancellation after partial verified MOVE cleanup emits `bridge.move.cleanup.cancelled`. |
| DR-16 | ✅ Fixed | Conflict prompt naming covers `Exists`, `NonEmptyDirectory`, and `ReparsePoint` buckets for child collisions. |

---

## Findings (ranked most-severe first)

### DR-1 — CRITICAL · Cross-FS MOVE source cleanup follows directory junctions and deletes files *outside* the moved tree
**File:** [FolderWindow.FileOperations.State.cpp:7160](RedSalamander/FolderWindow.FileOperations.State.cpp:7160) (`DeleteCopiedSourceEntryForMove`)
**Independently re-confirmed against live source.**

`DeleteCopiedSourceEntryForMove` computes `sourceIsDirectory` from `FILE_ATTRIBUTE_DIRECTORY` only (line 7160) — **no `FILE_ATTRIBUTE_REPARSE_POINT` check**. For a directory it calls `sourceFs.ReadDirectoryInfo` (line 7180), which *follows* a junction, then recurses (line 7229) and deletes each child via `DeleteCopiedSourcePathNoRecursive` → `sourceFs.DeleteItem`.

Confirmed reachability chain:
- `ReparsePointPolicy::FollowTargets` is a real, parseable policy ([State.cpp:264-299](RedSalamander/FolderWindow.FileOperations.State.cpp:264), from FS capability JSON or settings).
- Under `FollowTargets` the bridge copy *descends through* a junction `C:\src\link → D:\realdata`, copying `D:\realdata`'s files and **recording them in the manifest** keyed by through-junction source paths.
- `ShouldDeleteMoveSourceAfterBridgeCopy` ([State.cpp:8787](RedSalamander/FolderWindow.FileOperations.State.cpp:8787)) gates only on *skipped* reparse points — which `FollowTargets` never produces — so the cleanup runs.
- Cleanup recurses into `C:\src\link` as an ordinary directory, `ReadDirectoryInfo` follows it, and each through-junction child (present in the manifest) is `DeleteItem`-ed → **`D:\realdata`'s files are permanently deleted**, even though the user moved only `C:\src`.

The pre-delta code used a single `DeleteItem(... | FILESYSTEM_FLAG_RECURSIVE)`, and the local FS recursive delete explicitly does **not** descend into reparse points ([FileSystem.DirectoryOps.cpp:830](Plugins/FileSystem/FileSystem.DirectoryOps.cpp:830)) — so this is a **new** data-loss vector introduced by `a37fdde7f`.

**Fix:** after fetching `sourceAttributes`, add `const bool sourceIsReparse = (sourceAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0;`. When a directory source is a reparse point, do **not** enumerate/recurse — either `DeleteCopiedSourcePathNoRecursive` the junction itself (unlinks without following) or `NoteMoveSourceCleanupSkipped` and leave it. Mirror the FS guard at `FileSystem.DirectoryOps.cpp:830`. Add a self-test: move a tree containing a junction under `FollowTargets`, assert the junction target's external files survive.

### DR-2 — HIGH · Unconditional post-commit `SetFileAttributesW` reports a fully-promoted overwrite as failed
**File:** [FileSystem.Path.cpp:428](Plugins/FileSystem/FileSystem.Path.cpp:428) (`PromoteStagedTempIntoFinalPath`)

After `ReplaceFileW`/`MoveFileExW` succeed and `committed = true` (line 421), the post-commit `SetFileAttributesW(finalPath, desiredFinalAttributes)` (line 428) runs **unconditionally** (unlike the guarded temp-path call at line 351). If that metadata write is rejected (transient AV/sharing on the just-renamed handle, EFS/cloud placeholder provider, network redirector denying `FILE_WRITE_ATTRIBUTES`), line 431 returns a FAILED HRESULT for a file that is **already promoted**. Propagation: the writer path returns failure without `_committed` and its destructor deletes the consumed temp while the destination is silently overwritten; the copy/move engine reports the copy as failed and a MOVE leaves the source behind as a duplicate.

**Fix:** once `committed`, treat the attribute set as best-effort — only call it when `desiredFinalAttributes` differs from current, and swallow failure (`static_cast<void>(::SetFileAttributesW(...))`, optionally `Debug::Warning`). Return `S_OK` once committed.

### DR-3 — HIGH (regression) · Parallel name-only search swallows mid-enumeration errors and reports SUCCESS with truncated results
**File:** [FileSystem.Search.cpp:608](Plugins/FileSystem/FileSystem.Search.cpp:608) (`HasQueuedParallelDirectoryResultWork` live line ~582; reset at ~608)

The FV-search fix changed `HasQueuedParallelDirectoryResultWork` to `return !result.matches.empty();`, dropping the `SUCCEEDED(result.status)` gate. Now when a directory accumulates ≥1 match and then `IFilesInformation::Get` fails mid-enumeration (status set + break), the unconditional final `QueueParallelDirectoryResultChunk` no longer early-returns, moves matches out, and **resets `result.status = S_OK`**. The terminal check (`FAILED(result.status) && != kCancelledHr`) then reads S_OK, so `state.terminalHr` stays S_OK and the consumer reports **success with silently truncated results**. Pre-fix, a failing-status result with matches made the predicate return false, preserving the error.

**Fix:** snapshot `const HRESULT pendingStatus = result.status;` before the final flush and restore it after (matches already moved out), or only reset to S_OK in the helper when `SUCCEEDED(result.status)`. Add a self-test injecting an `IFilesInformation::Get` failure after ≥1 match and asserting the failure HR surfaces via `terminalHr`.

### DR-4 — HIGH · Cross-FS MOVE cleanup re-reads/re-downloads full file content to verify before deleting source (≈3× I/O)
**File:** [FolderWindow.FileOperations.State.cpp:7108](RedSalamander/FolderWindow.FileOperations.State.cpp:7108) (`CopiedFileStillMatchesDestination` → `ReadersHaveEqualContent`)

For every cross-FS MOVE, cleanup reads the **entire** file from both the source (re-read) and the destination (re-download) and `memcmp`s before deleting the source — unconditionally, with no settings gate. Moving 50 GB to S3/SFTP/curl/cloud becomes ~50 GB write + ~50 GB local re-read + ~50 GB remote **re-download** ≈ 150 GB I/O, re-downloading the just-uploaded payload over a metered/slow link. With 100k small files it also explodes round-trips, run serially (ignores `copyMoveMaxConcurrency`).

**Fix:** gate the byte-for-byte re-read behind an opt-in setting (default OFF for remote/cross-FS); rely on `BasicFileInfoStillMatchesSource` (creation/last-write/attributes) + size equality to confirm before delete. Never re-download the destination for remote filesystems — trust the bridge copy's byte-count. Optionally parallelize verification under the existing limiter.

### DR-5 — HIGH · `ReadersHaveEqualContent` requires equal per-chunk read sizes, falsely failing valid cross-FS moves
**File:** [FolderWindow.FileOperations.State.cpp:7047](RedSalamander/FolderWindow.FileOperations.State.cpp:7047)

The `IFileReader` contract ([FileSystem.h:534](Common/PlugInterfaces/FileSystem.h:534)) permits short non-zero reads mid-file; only `bytesRead==0` means EOF. `ReadersHaveEqualContent` compares `sourceRead != destinationRead || sourceRead == 0 || memcmp(...)` per chunk. When the two readers chunk byte-identical content differently — e.g. local NTFS returns a full 1 MiB while the `FileSystemCurl` streaming reader returns `min(bytesToRead, _bufferedBytes)` ([FileSystemCurl.DirectoryOps.cpp:537](Plugins/FileSystemCurl/FileSystemCurl.DirectoryOps.cpp:537)) — the first iteration sees `1048576 != 262144` and declares inequality on byte-correct data. Result: a successful move is reported PARTIAL and the source is left behind (move silently degrades to copy).

**Fix:** don't require per-chunk lengths to match. Refill each side's buffer with an inner read-until-full-or-EOF loop, compare `min(filled)` bytes, advance both by that amount, re-read only the unconsumed tail; declare inequality only at a real EOF-length mismatch or a `memcmp` difference over the common prefix.

### DR-6 — MEDIUM · Retained `moveBridge` captures a dangling per-iteration `cookie` pointer
**File:** [FolderWindow.FileOperations.State.cpp:9015](RedSalamander/FolderWindow.FileOperations.State.cpp:9015) (mirror at ~9515)

`moveBridge` is declared outside the per-item `for(;;)` but `cookie` is a per-iteration local; the bridge captures `&cookie` of the iteration that copied. If `DeleteCopiedSourceForMove` fails and a user Retry/Overwrite hits `continue`, the next iteration re-runs cleanup via the retained `moveBridge` whose `cookie` pointer now targets the destroyed local; `DeleteItem(..., cookie)` dereferences/writes it (`issueRetryCounts`). In practice MSVC reuses the stack slot (a fresh live `cookie` sits there — well-defined under `[basic.life]`, so the visible effect is reset bookkeeping), but a hard UAF arises if the compiler ever relocates the local. Pre-delta passed `nullptr`, so no dangling pointer existed.

**Fix:** match the cookie's lifetime to `moveBridge` — hoist `PerItemCallbackCookie` out of the loop (reset per iteration), re-point `moveBridge->cookie` at each retry, or construct the bridge with `cookie==nullptr` for the cleanup path. Apply to both loops.

### DR-7 — MEDIUM · Transient network probe failures during cleanup discard a verified-good copy as `ERROR_PARTIAL_COPY`
**File:** [FolderWindow.FileOperations.State.cpp:7321](RedSalamander/FolderWindow.FileOperations.State.cpp:7321)

After a fully successful cross-FS copy, cleanup re-probes source and destination; every non-cancellation, non-`IsMissingPathHr` failure is swallowed into `stats.anySkipped` → `ERROR_PARTIAL_COPY`. A network blip during cleanup (`GetAttributes`/`ReadDirectoryInfo`/`GetSize`/`Read` on either endpoint) converts a byte-correct copy into PARTIAL: destination not rolled back, source not deleted → **fully duplicated data** plus a PARTIAL error forcing a manual re-run.

**Fix:** distinguish transient/retryable probe failures from "source changed / cannot verify". Bound the re-probes with short retry/backoff on transient HRESULTs (mirroring the copy-phase circuit breaker); track transient-skip separately and surface it as a *retryable* cleanup error (retry just the delete), keeping keep-source only for confirmed mismatches.

### DR-8 — MEDIUM · `ReplaceFileW` lost `REPLACEFILE_IGNORE_MERGE_ERRORS` → destination ADS/owner/SACL dropped via MoveFileExW fallback
**File:** [FileSystem.Path.cpp:404](Plugins/FileSystem/FileSystem.Path.cpp:404)

The consolidated promoter passes `0u` for `dwReplaceFlags`; the prior writer path passed `REPLACEFILE_IGNORE_MERGE_ERRORS`. Overwriting a destination with named alternate data streams (Zone.Identifier/MotW or custom ADS) whose ACL/security merge produces a **non-fatal** merge error (network share, EFS dest) now makes `ReplaceFileW` return 0, routing to the `MoveFileExW` fallback which replaces the destination wholesale with the temp (no ADS). `RestoreFileMetadataAfterFallback` restores only DACL + times — so destination ADS/SACL/owner/object-id/short-name are silently lost. Affects both writer-overwrite and cross-FS move paths.

**Fix:** add `bool ignoreReplaceMergeErrors` to `StagedPromotionOptions`, set true from the writer caller (and consider always), pass `options.ignoreReplaceMergeErrors ? REPLACEFILE_IGNORE_MERGE_ERRORS : 0u`.

### DR-9 — MEDIUM · MoveFileExW fallback stamps the OLD destination's last-write time onto fresh content
**File:** [FileSystem.Path.cpp:425](Plugins/FileSystem/FileSystem.Path.cpp:425)

`CaptureFileMetadataForFallback` snapshots the *old* destination's times; on the `MoveFileExW` fallback, `RestoreFileMetadataAfterFallback` calls `SetFileTime` with all three, stamping the new content with the old `lastWriteTime`/`lastAccessTime` instead of the source's. The `ReplaceFileW` path yields the source's last-write time, so the fallback diverges: copying `newer-source` (mtime 2026-06-27) over `older-dest` (2020-01-01) yields 2020-01-01 — `robocopy /MIR`, rsync mirrors, and the project's own CompareDirectories last-write comparison then treat the fresh copy as unchanged/older.

**Fix:** restore only the destination's `creationTime` (and security); pass `nullptr` for last-write/last-access so the move-fallback keeps the source-derived time CopyFileExW already stamped on the temp. Add a self-test forcing a `ReplaceFileW` failure and asserting the promoted file's last-write equals the source's.

### DR-10 — MEDIUM · UNC share-root guard bypassed for forward-slash bare-share-root destinations
**File:** [FolderWindow.FileOperations.State.cpp:137](RedSalamander/FolderWindow.FileOperations.State.cpp:137) (`IsUncShareRootBoundary`)

`TryGetUncShareRootBoundary` normalizes its result to backslashes, but `IsUncShareRootBoundary` compares the **raw** path argument via `EqualsNoCase`. If a destination ever carries forward slashes and is exactly the bare share root (`//server/share`), the boundary (`\\server\share`) won't match, the guard is bypassed, and `GetAttributes`/`CreateDirectory` runs on the bare share root — the exact action R6 was meant to prevent. Latent today (native enumeration is backslash-form); also affects the pre-existing guard, so not newly introduced.

**Fix:** make the matcher separator-agnostic — normalize the incoming path to backslashes inside `IsUncShareRootBoundary` before `EqualsNoCase`, so parser and matcher agree.

### DR-11 — LOW · Overwrite symlink onto a non-reparse volume surfaces raw `ERROR_INVALID_PARAMETER` instead of `ERROR_NOT_SUPPORTED`
**File:** [FileSystem.FileOps.cpp:3948](Plugins/FileSystem/FileSystem.FileOps.cpp:3948)

The N4 remap was gated on `!allowOverwriteEffective`, so for the overwrite case a failed `COPY_FILE_COPY_SYMLINK` on FAT/exFAT/some SMB surfaces opaque `0x80070057` while the non-overwrite path still remaps to `ERROR_NOT_SUPPORTED`. Diagnostic-quality regression only (destination untouched, temp cleaned).

**Fix:** drop the `!allowOverwriteEffective &&` qualifier so the remap runs whenever `copyHr == ERROR_INVALID_PARAMETER` after a symlink copy; if promote-phase `ERROR_INVALID_PARAMETER` should stay raw, return a distinct copy-phase sentinel from the staged helper instead of gating on `allowOverwrite`.

### DR-12 — LOW · `BestEffortDeleteCopiedTemp` follows a file-symlink-to-directory and leaks the staged temp (old N6 unfixed)
**File:** [FileSystem.FileOps.cpp:2819](Plugins/FileSystem/FileSystem.FileOps.cpp:2819)

For a staged copy of a file symlink whose target is a directory, `GetFileAttributesW` follows the link, returns `FILE_ATTRIBUTE_DIRECTORY`, and `RemoveDirectoryW` is called on a file-type reparse point (refuses) with no `DeleteFileW` fallback — leaking the `.rs_copy_tmp_*` sibling (indexer skip only hides it). This is exactly old-N6, still unfixed.

**Fix:** detect `FILE_ATTRIBUTE_REPARSE_POINT` (or open with `FILE_FLAG_OPEN_REPARSE_POINT`) before treating the temp as a directory and use `DeleteFileW`; at minimum fall back to `DeleteFileW` when `RemoveDirectoryW` fails.

### DR-13 — LOW · `IsRedSalamanderStagedTempName` substring match hides real user files from search/index
**File:** [LocalSearchIndexCore.cpp:483](Common/LocalSearchIndexCore.cpp:483)

Unanchored `find()!=npos` for `.rs_copy_tmp_` / `.~rs-write-` means a real user file embedding either substring (e.g. `notes.rs_copy_tmp_backup.txt`) is silently excluded from enumeration, search, and the index. Genuine temps always carry the marker as a true suffix + hex/GUID tail, so a stricter check is feasible.

**Fix:** require the marker followed by the expected tail (e.g. `.rs_copy_tmp_` then `^[0-9A-Fa-f]{8}_[0-9A-Fa-f]{8}_[0-9A-Fa-f]{16}$`; `.~rs-write-` then a GUID/hex id ending in `.tmp`).

### DR-14 — LOW · Sequential staged-overwrite copy failure does not roll back `completedBytes`, inflating progress
**File:** [FileSystem.FileOps.cpp:3801](Plugins/FileSystem/FileSystem.FileOps.cpp:3801)

In a sequential recursive copy (`context.parallel==nullptr`), a child's staged-overwrite temp copy succeeds (crediting bytes via the progress routine) then `PromoteStagedTempIntoFinalPath` fails; `RollBackParallelStagedProgress` early-returns (parallel is null) and `returnFailure` never restores `context.completedBytes`. With continue-on-error/Skip, subsequent items carry the phantom bytes, permanently over-reporting (compounding per failure). Progress only; data integrity unaffected.

**Fix:** in the sequential failure path restore `context.completedBytes = progress.itemBaseBytes` (under the progress mutex), mirroring the parallel rollback.

### DR-15 — LOW · Cross-FS MOVE cancel mid-cleanup leaves a half-emptied source with no diagnostic
**File:** [FolderWindow.FileOperations.State.cpp:7232](RedSalamander/FolderWindow.FileOperations.State.cpp:7232)

Cancelling after K of N children are deleted returns `ERROR_CANCELLED` immediately, skipping the rest and the parent delete, and the cancellation short-circuit returns without logging or setting `anySkipped`. Not data loss (each deleted child has a verified destination copy), but the source is left half-emptied with no record.

**Fix:** emit a one-shot `bridge.move.cleanup.cancelled` diagnostic before unwinding, or surface a `cancelledPartial` flag so the caller records the partial-source condition.

### DR-16 — LOW · Merged-folder child conflict prompt no longer names the offending child for the new buckets
**File:** [FolderWindow.FileOperations.Popup.cpp:5269](RedSalamander/FolderWindow.FileOperations.Popup.cpp:5269)

After the conflict-bucket split, a child collision surfacing as `ERROR_DIR_NOT_EMPTY`/`ERROR_REPARSE_POINT_ENCOUNTERED` is now `NonEmptyDirectory`/`ReparsePoint` instead of `Exists`. The named-leaf message branch is gated solely on `bucket==Exists`, so the new buckets fall through to a generic message that no longer names the offending child (cosmetic/UX; From/To and the Overwrite action remain correct).

**Fix:** generalize the named-leaf gate to also cover `NonEmptyDirectory`/`ReparsePoint` (add `*_NAMED` strings to `RedSalamander.rc` + satellites, or reuse one named phrasing).

---

## Suggested order of attack

1. **DR-1 (critical)** — add the reparse guard to `DeleteCopiedSourceEntryForMove` before anything else; this can delete files outside the moved set. Add the junction self-test.
2. **DR-3 (regression)** + **DR-2** — both surface as wrong success/failure signals; small, contained fixes in `FileSystem.Search.cpp` and `FileSystem.Path.cpp`.
3. **DR-5** then **DR-4** — the move-verification correctness bug (false PARTIAL on streaming readers) then the ~3× I/O strategy (remote-aware fast path / opt-in).
4. **DR-6, DR-7** — move-cleanup lifetime + transient-failure robustness.
5. **DR-8, DR-9, DR-10** — promoter metadata fidelity + UNC matcher symmetry.
6. **DR-11 … DR-16** — diagnostics/UX/progress cleanups.

## Verified-correct (no action)
FV-retrycap (F1 retry caps), FV-small (F4/F8/F10). The unified `PromoteStagedTempIntoFinalPath` correctly fixes the original F5/N1/N5 readonly + ACL/time-on-fallback + commit issues (the residuals above are *new* sub-cases, not the originals).
