# Operation FS Deep Audit Track B - Cross-FS Directory Move Delete Safety - 2026-06-26

## Status

- State: Complete; moved to `Specs/Plans/Done/`.
- Priority: P0 data safety.
- Scope: Cross-filesystem directory MOVE source cleanup after copy.

## Problem

Cross-filesystem directory MOVE must not recursively delete a live source tree from a stale enumeration snapshot. Entries created or modified after the copy snapshot must be preserved unless the destination is proven to contain the same version.

## Targets

- `RedSalamander/FolderWindow.FileOperations.State.cpp`
- `Plugins/FileSystem/FileSystem.FileOps.cpp`
- `RedSalamander/SelfTest/FileOperations/*`
- `Specs/FileSystem/FileSystem_FileOperations.md`

## Tasks

1. [x] Define cross-FS directory MOVE cleanup as per-entry cleanup, not blanket recursive delete.
2. [x] Record a copied-entry manifest with relative path, kind, copied size, and stable metadata available through the bridge (`creationTime`, `lastWriteTime`, and attributes). Provider identity/version/etag is not currently exposed by the bridge contract, so it remains unavailable for this implementation.
3. [x] Delete only source entries that are proven copied and unchanged.
4. [x] Preserve files created after the copy snapshot; preserve and report partial for modified files.
5. [x] Return `ERROR_PARTIAL_COPY` and surface diagnostics when source cleanup is intentionally incomplete.
6. [x] Add deterministic FileOps tests for added-after-snapshot, modified-after-copy, and extra-child-prevents-directory-removal behavior (`Floodgate_CrossFsDirectoryMoveCleanupPreservesChangedSource`). Reparse no-follow remains covered by the local filesystem fallback semantics and existing bridge policy; no additional host-bridge reparse regression was needed for this Track B fix.
7. [x] Update `Specs/FileSystem/FileSystem_FileOperations.md`.

## Implementation Notes

- `CrossFileSystemBridge` now records copied directories and files in a copied-entry manifest as each destination entry is created/promoted.
- Cross-filesystem Move source cleanup now walks that manifest and deletes source files only after validating that the source entry still matches the copied/promoted destination content and stable metadata.
- Source deletes during Move cleanup are non-recursive. Newly created children or changed descendants keep the source directory alive and cause the item to finish as partial cleanup instead of deleting live data from a stale snapshot.
- The concurrent and serial cross-filesystem Move paths keep the bridge state alive through source cleanup so cleanup uses the manifest from the copy phase.
- Test-only pause hooks create a deterministic race window between copy completion and source cleanup for regression coverage.
- 2026-06-28 revalidation found one remaining real cleanup bug: when a non-reparse source directory's own metadata changed after copy, cleanup preserved the directory but also skipped all copied children under it. Cleanup now still enumerates that changed directory, deletes copied descendants that remain proven unchanged, and preserves only the changed directory itself. Changed reparse directories still fail closed and are not enumerated.

## Validation

```powershell
.\build.ps1 -ProjectName RedSalamander
# PASS, 0 errors, 0 warnings; log: .build\logs\msbuild-20260627_100341_653.log

.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Floodgate_CrossFsDirectoryMoveCleanupPreservesChangedSource
# PASS, 3 passed, 0 failed

.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Floodgate_CrossFs
# PASS, 5 passed, 0 failed

.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Fairstream_CrossFsConcurrentMoveUsesBridge
# PASS, 3 passed, 0 failed

.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Riptide_Bridge
# PASS, 6 passed, 0 failed

.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter FileOpsFamily_Fairstream
# PASS, 39 passed, 0 failed
```

`git diff --check`:

- PASS; no whitespace errors. Git reported line-ending normalization warnings only.

Evidence archive:

- `Specs/TestRuns/SINON/FileOps/2026-06-27_TrackB_CrossFsDirectoryMoveDeleteSafety/`
- 2026-06-28 focused revalidation: `Floodgate_CrossFsDirectoryMoveCleanupPreservesChangedSource` passed (`3 passed / 0 failed`).
- 2026-06-28 full FileOps revalidation: `101 passed / 0 failed / 20 skipped`. Evidence: `Specs/TestRuns/7d3a1247382a/FileOps/2026-06-28_152421/fileops_results.json`.
