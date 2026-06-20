# Operation Keystone — 3-Day Review Remediation (2026-06-20)

Tracking plan for findings from the multi-agent review of the last 3 days of work
(diff `a68274ade..804309513`: Truename batch-rename remediation, the new `FileSystemPathIdentity`
module, cloud copy/move/delete hardening, and FolderView perf instrumentation).

Two adversarially-verified review rounds (round 1: 32→29 confirmed; round 2 gap review: 13→9 confirmed)
plus a completeness critic. Each item below is **confirmed** (survived two independent skeptics) unless
noted. Severity is the verifier-adjusted severity.

## Already fixed in the 2026-06-20 session (built clean, see commit/working tree)

- **[HIGH] Curl staged re-stat broke copy/MOVE to sizeless-`LIST` FTP servers** — added `Entry::sizeKnown`;
  staged size check only enforced when known (`FileSystemCurl.*`). Residual proof-depth tracked under
  `Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md` (FG-P0-3-FALLBACK-SIZE-ZERO).
- **[HIGH] MSDrive merge-move recursive source delete without re-confirming empty** — now re-lists the
  source folder (fresh, complete, empty) before `DeleteItemById`; otherwise preserves source + reports
  partial (`FileSystemMicrosoftDrive.cpp` `MergeMoveFolderIntoExisting`). **Needs a regression selftest**
  (fake Graph: child omitted from first snapshot then present on re-list → source preserved) — see KS-T1.
- **[HIGH] PluginContractTests did not assert the mandatory `pathIdentity` block parses** — added a
  Step-3 assertion calling the exact host parser `TryParseFileSystemPathIdentityContract` for every
  shipped plugin (closes the fail-closed bricking gate). All 7 plugins green.
- **[perf] JSONL perf sink** — resolve env once + atomic fast-path so every `Debug::Perf::Emit` is free
  in production (`PerfJsonl.cpp`); gated the one ungated layout-pass emit (`FolderView.Layout.cpp`).
- **[bug] BatchRename cycle-break temp leaf** length-capped to 255 (`BatchRenameExecutionEngine.cpp`).
- **[cleanup] BatchRename** folded the duplicate perf-emit gate; **doc** `PluginHostModel.md` updated for
  the `pathIdentity` fail-closed gate.

## Routed — correctness / data-safety (design or test work, not safe to hot-fix mid-review)

| ID | Sev | Item | Fix direction | Proof |
|----|-----|------|---------------|-------|
| KS-1 | **HIGH** | Cross-plugin MOVE verification is inconsistent: **S3 `PerformTransfer` does NO post-copy verification** before `DeleteS3Object` (`FileSystemS3.Directory.cpp:1686`); the **cross-FS bridge verifies size-only** before source delete (`State.cpp:7386`); **Curl** falls back to existence-only on sizeless servers; only **local** does a full byte-compare. The least-reliable media have the weakest pre-delete checks. | Define one host-level "MOVE never deletes a source it hasn't positively verified arrived intact" contract. Per backend, use the strongest cheap check: S3 ETag/`ChecksumSHA256` on PUT/Copy vs source; bridge rolling hash computed while streaming vs a post-promote read; Curl `SIZE`/`HASH`/`XCRC` (or don't delete when unverifiable). MSDrive exempt (atomic server-side PATCH, no unverified delete window). | Fault-injected same-length-corrupt destination on each backend → MOVE returns `ERROR_PARTIAL_COPY`, source preserved. |
| KS-2 | MEDIUM | **Curl `caseOnlyRename:"supported"` + case-sensitive identity, but rename skips case-only renames as case-INSENSITIVE self-renames** → silent success, file unchanged (`RenameItem` 2552, `MoveItem` 2290, `MoveItems` 3081, `RenameItems` 3650; `EqualsInsensitive` 390). | Detect self-renames case-SENSITIVELY for case-sensitive providers (exact `remotePath ==`), or drive off `pathIdentity.componentComparison`. If a protocol can't honor it, set `caseOnlyRename:"unsupported"`/`"noOp"`. | Case-only rename over FTP performs the rename (or returns a real error), never false success. |
| KS-3 | MEDIUM | **MSDrive `caseOnlyRename:"supported"` but `MoveOrRenameItem` returns `S_OK` when destination resolves to the same item id** (always true for case-only on case-insensitive OneDrive) (`FileSystemMicrosoftDrive.cpp:4168`). _(split vote +1/-1; verify against Graph behavior.)_ | If same-id and leaves differ only by case, issue the Graph name PATCH; else set `caseOnlyRename:"noOp"`. | Case-only rename changes the name, or capability honestly reports no-op. |
| KS-4 | MEDIUM | **`EnsureDestinationDirectory` create/stat race** falls back to `ERROR_ALREADY_EXISTS` after 3 attempts, aborting the (sub)tree instead of treating an existing directory as success (`State.cpp:~1838`). Critic flagged possible partial-move escalation. | Treat "exists as a directory" as success; trace whether sibling source deletion can already have run when a subtree aborts. | Concurrent-create race resolves as success; no orphaned/under-copied destination subtree. |
| KS-5 | MEDIUM | **Plugin debug self-tests silently pass when the `_DEBUG`-gated export is absent** (`PluginContractTests.cpp` `TestPluginDebugSelfTests`, `pfnSelfTests==nullptr` branch). Same "green test exercised nothing" class as the FG-P0-2 trap. | Fail (or explicitly xfail) when a plugin that should ship `RunDebugSelfTests` does not export it. | Removing an export turns the step RED. |
| KS-6 | LOW | **Copy onto a byte-identical existing file silently returns `S_OK`** without prompting; destination keeps its own timestamps/ACLs/ADS (`FileSystem.FileOps.cpp:~3518`). Data-safe but changes semantics + leaves metadata divergent. | Confirm intended for non-resume copies; consider reporting the skip to the host or syncing source metadata. | Documented, or host-visible skip. |
| KS-7 | LOW | **FindFiles partial move/delete keeps ALL result rows** (phantom stale rows for items whose source was removed) (`FindFilesWindow.cpp` `OnFolderWindowFileOperationCompleted`). Fail-safe direction (never hides live rows). | Surface per-source results in the completed event; remove only completed keys. | Partial move removes only moved rows. |
| KS-8 | LOW | **Curl directory MOVE deletes the whole source subtree without per-child verification** (pre-existing, `MoveItem` dir branch ~2338). | Fold into the KS-1 verify-before-delete contract for recursive cloud moves. | Per-child verification before recursive source delete. |

## Routed — performance ("the code must be fast")

| ID | Sev | Item | Fix direction |
|----|-----|------|---------------|
| KS-P1 | MEDIUM | **BatchRename per-keystroke preview does N synchronous FS stat calls on the UI thread** for local batches (`BatchRenameWindow.cpp` `RebuildPreview`→`ApplyLocalDestinationConflictValidation`). | Debounce + move conflict stat off the UI thread (or cache dir listing per refresh). |
| KS-P2 | MEDIUM | **Provider selection collection runs network/archive I/O synchronously on the UI thread** in `SetContext`/`Create` (`CollectProviderSelectionTargets`). | Collect targets off-thread; show progress. |
| KS-P3 | MEDIUM | **S3 per-object destination pre-pass duplicates the ancestor-walk HEAD probes** done again in the transfer loop (`ExecuteCopyOrMove` ~1432 vs ~1486). | Reuse the pre-pass result; one HEAD per object. |
| KS-P4 | MEDIUM | **A single non-ASCII filename forces whole-batch O(n²) duplicate detection** because the key builder bails on any char >0x7F (`FileSystemPathIdentity.cpp` `TryAppendKeyChar`; `BatchRenameEngine.cpp` `MarkDuplicateTargets`). Common for real filenames. | Fold non-ASCII case in the key builder (e.g. `LCMapStringEx` consistent with `CompareStringOrdinal`) so the O(n) hashed path is retained. Ties to KS-A1. |
| KS-P5 | LOW | Cross-volume MOVE re-reads every byte before delete (read I/O doubled) — this is the *safe* content verify; reconcile with KS-1 (reuse a hash computed during copy). |
| KS-P6 | LOW | `HasPlannedDestinationAncestorCollision` allocates a `std::string` per path segment per object (`FileSystemS3.Directory.cpp:~1171`); storage-medium probe issues volume IOCTLs on the concurrency-resolution path (`FileSystem.cpp` `ProbeLocalStorageMedium`). |

## Routed — simplicity / architecture ("the simplest possible")

| ID | Sev | Item | Fix direction |
|----|-----|------|---------------|
| KS-S1 | MEDIUM | **`GenerateRandomBytes` + `AppendHexToken` reimplemented independently in 3 plugins** (S3, Curl, MSDrive). | One shared crypto-random/hex helper in Common. |
| KS-S2 | MEDIUM | **Four divergent "generate a unique temp/sibling name" schemes** for the same data-safety purpose (S3 `BuildHiddenSiblingKey`, Curl `BuildRemoteSiblingLeaf`, BatchRename `MakeBatchRenameTempLeaf`, …). | Unify on one helper (with the 255-cap now in `MakeBatchRenameTempLeaf`). |
| KS-S3 | LOW | **`DebugCheck` self-test harness copy-pasted into 4 plugin TUs**; the large fake-S3 `DebugS3Graph` harness + `thread_local` global + redirect shims live in a **production** TU (`FileSystemS3.Directory.cpp`). | Extract shared test harness to a `_DEBUG`/test-only header or TU. |
| KS-S4 | LOW | BatchRename re-derives separator/prefix logic locally (`GuessPreferredSeparator`, `JoinFolderAndLeaf`) and mixes two path-decomposition schemes (manual accepted-separator split vs `std::filesystem::path` iteration) in `ApplyExecutedDirectoryMoves` — instead of using the new `FileSystemPathIdentity`. | Route all path identity/decomposition through `FileSystemPathIdentity`. |
| KS-S5 | LOW | Dead `SkipAll` handling scattered across the file-op state machine after the prompt action was removed (`State.cpp` `BuildConflictActionLayout` etc.). | Remove the dead branches. |
| KS-S6 | LOW | `MakeDebugFileSystem` uses raw `new` (house-rule deviation, debug-only) (`FileSystemMicrosoftDrive.cpp`); test-only manual `DestroyWindow` after `release()` in `State.Queue.Part.cpp`. | Use smart pointers / WIL even in debug scaffolding. |
| KS-A1 | MEDIUM | **PathIdentity has two equivalence relations**: `EquivalentComponent` folds full Unicode via `CompareStringOrdinal`, but the key builders fold ASCII-only and decline (`nullopt`) on >0x7F. Correct today only by the decline-and-fallback convention (no live bug — callers honor it). | Document the invariant in `FileSystemPathIdentity.h` (nullopt = "must fall back to EquivalentPath", not "no match"); ideally fold non-ASCII in the key builders too (also fixes KS-P4). |
| KS-A2 | LOW | `EquivalentComponent`'s `CompareStringOrdinal` case-fold may disagree with the NTFS `$UpCase` table for rare code points. | Document the approximation, or existence-check before treating two names as the same object. |

## Tests to add (lock the fixes / close the recurring "green test exercised nothing" class)

| ID | For | Required proof |
|----|-----|----------------|
| KS-T1 | MSDrive merge-move fix | Fake Graph: a child omitted from the first enumeration snapshot but present on the pre-delete re-list → source folder preserved, op returns partial. Removing the re-list makes it RED. |
| KS-T2 | KS-1 | Per-backend same-length-corrupt fault injection (see KS-1 proof). |
| KS-T3 | Systemic | Build-time/Pester guard: every non-Setup/Cleanup `kFileOpsPhaseOrder` case is a member of exactly one `kFileOpsFamily*` array (mirrors the FG-HARNESS-COVERAGE item; kills the silent-skip class for good). |

## Refuted (do NOT re-flag)

- `EquivalentPath` "doesn't collapse doubled/trailing separators → self-overwrite" — **refuted both rounds**
  (inputs are normalized upstream; no path reaches a self-overwrite).
- `pathTextStableIdentity:false` "loses all duplicate/cycle protection" — **refuted**: the collapse is
  fail-closed and properly gated; no unprotected rename/move reaches execution (GoogleDrive ops disabled).
- Capability-parse failure "silently refuses with no user-visible error" — refuted (medium).
- Reparse-point source link "deleted on mere destination existence" — refuted.
