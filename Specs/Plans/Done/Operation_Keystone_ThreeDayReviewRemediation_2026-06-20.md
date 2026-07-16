# Operation Keystone — 3-Day Review Remediation (2026-06-20)

> **2026-07-02 folder review:** ledger verified against HEAD `275c04034` — rows below are all still
> open; superseded rows moved to the routed list.

Tracking plan for findings from the multi-agent review of the last 3 days of work
(diff `a68274ade..804309513`: Truename batch-rename remediation, the new `FileSystemPathIdentity`
module, cloud copy/move/delete hardening, and FolderView perf instrumentation).

Two adversarially-verified review rounds (round 1: 32→29 confirmed; round 2 gap review: 13→9 confirmed)
plus a completeness critic. Each item below is **confirmed** (survived two independent skeptics) unless
noted. Severity is the verifier-adjusted severity.

## Routed / superseded (2026-07-02 folder review)

- **2026-06-20 session fixes** (Curl sizeless-`LIST` re-stat, MSDrive merge-move pre-delete re-list
  guard, `pathIdentity` contract assertion in PluginContractTests, JSONL perf sink fast-path,
  BatchRename temp-leaf 255 cap, cleanups) — all verified landed, incl. **KS-T1**, which landed as
  `RunDebugReListGuardBlocksConcurrentChildSelfTest` in `385e5ba5a`.
- **KS-1 / KS-8 / KS-T2** (MOVE verify-before-delete contract + fault-injection tests) — provider
  legs adjudicated and pinned by Done FSDeepAudit TrackB/TrackF (2026-06-26/27); the cross-FS bridge
  leg is owned by Clearwater CW-P1/CW-T-P1.
- **KS-4** — stale premise: `EnsureDestinationDirectory` already re-stats and returns `S_OK` when the
  destination exists as a directory; the race hook predates the spec.
- **KS-P3** — duplicate of Floodgate FG-P2-5 (single owner: Floodgate).
- **KS-P5** — duplicate of Clearwater CW-P1 (single owner: Clearwater).
- **KS-P6 (first half)** — done: `HasPlannedDestinationAncestorCollision` was replaced by the
  `plannedDestinationKeys` set. Second half (ProbeLocalStorageMedium IOCTL probe) kept as live row
  KS-P6 below.
- **KS-S6** — raw-`new` idiom adjudicated by Granite GR-C2 (new+attach is the house model); the
  residual `release()`+`DestroyWindow` at `State.Queue.Part.cpp:22-33` is a deliberate
  destroy-outside-lock pattern.
- **KS-T3** — duplicate of Floodgate FG-HARNESS-COVERAGE (single owner: Floodgate).

## Live — correctness / data-safety

**NEXT ACTION (2026-07-02 folder review): KS-2 + KS-3** — both providers still advertise
`caseOnlyRename:"supported"` while silently no-opping case-only renames; fix the rename paths or
honestly downgrade the capability.

| ID | Sev | Item | Fix direction | Proof |
|----|-----|------|---------------|-------|
| KS-2 | MEDIUM | **Curl `caseOnlyRename:"supported"` + case-sensitive identity, but rename skips case-only renames as case-INSENSITIVE self-renames** → silent success, file unchanged (`FileSystemCurl.CopyMove.cpp:2306/2568/3097/3666`; `EqualsInsensitive`). Verified still open 2026-07-02, capability still `"supported"`. | Detect self-renames case-SENSITIVELY for case-sensitive providers (exact `remotePath ==`), or drive off `pathIdentity.componentComparison`. If a protocol can't honor it, set `caseOnlyRename:"unsupported"`/`"noOp"`. | Case-only rename over FTP performs the rename (or returns a real error), never false success. |
| KS-3 | MEDIUM | **MSDrive `caseOnlyRename:"supported"` but `MoveOrRenameItem` returns `S_OK` when destination resolves to the same item id** (always true for case-only on case-insensitive OneDrive) (`FileSystemMicrosoftDrive.cpp:4155`). Verified still open 2026-07-02, capability still `"supported"`. _(split vote +1/-1; verify against Graph behavior.)_ | If same-id and leaves differ only by case, issue the Graph name PATCH; else set `caseOnlyRename:"noOp"`. | Case-only rename changes the name, or capability honestly reports no-op. |
| KS-5 | MEDIUM | **Plugin debug self-tests silently pass when the `_DEBUG`-gated export is absent** (`PluginContractTests.cpp:399-404` `TestPluginDebugSelfTests`, `pfnSelfTests==nullptr` branch). Same "green test exercised nothing" class as the FG-P0-2 trap. | Fail (or explicitly xfail) when a plugin that should ship `RunDebugSelfTests` does not export it. | Removing an export turns the step RED. |
| KS-6 | LOW | **Copy onto a byte-identical existing file silently returns `S_OK`** without prompting; destination keeps its own timestamps/ACLs/ADS (`FileSystem.FileOps.cpp:~3518`). Data-safe but changes semantics + leaves metadata divergent. 2026-07-02 folder review: LOW decision item. | Confirm intended for non-resume copies; consider reporting the skip to the host or syncing source metadata. | Documented, or host-visible skip. |
| KS-7 | LOW | **FindFiles partial move/delete keeps ALL result rows** (phantom stale rows for items whose source was removed) (`FindFilesWindow.cpp` `OnFolderWindowFileOperationCompleted`). 2026-07-02 folder review: deliberate fail-safe (never hides live rows); treat as enhancement. | Surface per-source results in the completed event; remove only completed keys. | Partial move removes only moved rows. |

## Live — performance ("the code must be fast")

2026-07-02 folder review: this plan is now the **single owner** of live batch-rename perf work
(KS-P1/KS-P2) since `UI_BatchRenameWindowPlan` moved to Done.

| ID | Sev | Item | Fix direction |
|----|-----|------|---------------|
| KS-P1 | MEDIUM | **BatchRename per-keystroke preview does N synchronous FS stat calls on the UI thread** for local batches (`BatchRenameWindow.cpp` `RebuildPreview`→`ApplyLocalDestinationConflictValidation`). | Debounce + move conflict stat off the UI thread (or cache dir listing per refresh). |
| KS-P2 | MEDIUM | **Provider selection collection runs network/archive I/O synchronously on the UI thread** in `SetContext`/`Create` (`CollectProviderSelectionTargets`). | Collect targets off-thread; show progress. |
| KS-P4 | MEDIUM | **A single non-ASCII filename forces whole-batch O(n²) duplicate detection** because the key builder bails on any char >0x7F (`FileSystemPathIdentity.cpp:181-183` `TryAppendKeyChar`; `BatchRenameEngine.cpp` `MarkDuplicateTargets`). Common for real filenames. | Fold non-ASCII case in the key builder (e.g. `LCMapStringEx` consistent with `CompareStringOrdinal`) so the O(n) hashed path is retained. Ties to KS-A1. |
| KS-P6 | LOW | Storage-medium probe issues volume IOCTLs on the concurrency-resolution path (`FileSystem.cpp` `ProbeLocalStorageMedium`). (First half — per-segment `std::string` allocations — done; see routed list.) | Cache/defer the IOCTL probe off the resolution path. |

## Live — simplicity / architecture ("the simplest possible")

| ID | Sev | Item | Fix direction |
|----|-----|------|---------------|
| KS-S1 | MEDIUM | **`GenerateRandomBytes` + `AppendHexToken` reimplemented independently in 3 plugins** (S3, Curl, MSDrive). | One shared crypto-random/hex helper in Common. |
| KS-S2 | MEDIUM | **Four divergent "generate a unique temp/sibling name" schemes** for the same data-safety purpose (S3 `BuildHiddenSiblingKey`, Curl `BuildRemoteSiblingLeaf`, BatchRename `MakeBatchRenameTempLeaf`, …). | Unify on one helper (with the 255-cap now in `MakeBatchRenameTempLeaf`). |
| KS-S3 | LOW | **`DebugCheck` self-test harness copy-pasted into 4 plugin TUs**; the large fake-S3 `DebugS3Graph` harness + `thread_local` global + redirect shims live in a **production** TU (`FileSystemS3.Directory.cpp`). | Extract shared test harness to a `_DEBUG`/test-only header or TU. |
| KS-S4 | LOW | BatchRename re-derives separator/prefix logic locally (`GuessPreferredSeparator`, `JoinFolderAndLeaf`) and mixes two path-decomposition schemes (manual accepted-separator split vs `std::filesystem::path` iteration) in `ApplyExecutedDirectoryMoves` — instead of using the new `FileSystemPathIdentity`. | Route all path identity/decomposition through `FileSystemPathIdentity`. |
| KS-S5 | LOW | Dead `SkipAll` handling scattered across the file-op state machine after the prompt action was removed (`State.cpp` `BuildConflictActionLayout` etc.). | Remove the dead branches. |
| KS-A1 | MEDIUM | **PathIdentity has two equivalence relations**: `EquivalentComponent` folds full Unicode via `CompareStringOrdinal`, but the key builders fold ASCII-only and decline (`nullopt`) on >0x7F. Correct today only by the decline-and-fallback convention (no live bug — callers honor it). | Document the invariant in `FileSystemPathIdentity.h` (nullopt = "must fall back to EquivalentPath", not "no match"); ideally fold non-ASCII in the key builders too (also fixes KS-P4). |
| KS-A2 | LOW | `EquivalentComponent`'s `CompareStringOrdinal` case-fold may disagree with the NTFS `$UpCase` table for rare code points. | Document the approximation, or existence-check before treating two names as the same object. |
