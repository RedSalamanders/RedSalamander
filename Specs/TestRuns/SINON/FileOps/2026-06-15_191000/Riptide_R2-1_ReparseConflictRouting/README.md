# Riptide R2-1 Reparse Conflict Routing

Scope: prove that directory reparse points cannot silently replace existing real directories without an overwrite grant, and that source reparse points do not enter the same-volume MOVE directory-merge fast path.

## RED

- Build: `.\build.ps1 -ProjectName RedSalamander -Configuration Debug`
  - Log: `RED-build-msbuild-20260615_185912_786.log`
  - Result: 0 warnings, 0 errors. This build included the new R2-1 selftests but not the production fix.
- Selftest: `Riptide_ReparseCopyOntoEmptyRealDirRequiresConsent`
  - Artifact: `RED_Riptide_ReparseCopyOntoEmptyRealDirRequiresConsent/`
  - Result: failed for the intended reason: the old engine completed `S_OK` where the test expected `ERROR_PARTIAL_COPY` after a conflict/Skip.
- Selftest: `Riptide_ReparseMoveSourceNeverUsesDirectoryMerge`
  - Artifact: `RED_Riptide_ReparseMoveSourceNeverUsesDirectoryMerge/`
  - Result: failed for the intended reason: the old engine completed `S_OK` where the test expected `ERROR_ALREADY_EXISTS`.

## GREEN

- Build after production fix: `.\build.ps1 -ProjectName RedSalamander -Configuration Debug`
  - Log: `GREEN-build-msbuild-20260615_190323_309.log`
  - Result: 0 warnings, 0 errors.
- Final build after updating the legacy Clearflow reparse test expectation:
  - Log: `GREEN-final-build-msbuild-20260615_190627_805.log`
  - Result: 0 warnings, 0 errors.
- Exact selftest: `Riptide_ReparseCopyOntoEmptyRealDirRequiresConsent`
  - Artifact: `GREEN_Riptide_ReparseCopyOntoEmptyRealDirRequiresConsent/`
  - Result: passed 3/0/0; confirms the empty real destination stays a real directory, one conflict is raised, sibling bytes are preserved, and the out-of-tree junction target sentinel is unchanged.
- Exact selftest: `Riptide_ReparseMoveSourceNeverUsesDirectoryMerge`
  - Artifact: `GREEN_Riptide_ReparseMoveSourceNeverUsesDirectoryMerge/`
  - Result: passed 3/0/0; confirms `followTargets` MOVE returns `ERROR_ALREADY_EXISTS` instead of moving data through the source junction target.
- Legacy regression: `FileOps_ReparseDirectoryMergeIntoExistingFolder`
  - Artifact: `GREEN_FileOps_ReparseDirectoryMergeIntoExistingFolder/`
  - Result: passed 3/0/0; now asserts no-grant rejection first, then explicit overwrite replacement.
- Policy regression: `Phase12_ReparsePointPolicy`
  - Artifact: `GREEN_Phase12_ReparsePointPolicy/`
  - Result: passed 3/0/0.
- Family regression: `FileOpsFamily_Fairstream`
  - Artifact: `GREEN_FileOpsFamily_Fairstream/`
  - Result: passed 28/0/0 with the two new R2-1 cases registered.

## Implementation Notes

- `CopyReparsePointInternal(...)` now mirrors the file-reparse branch: an existing directory without an overwrite grant returns `ERROR_ALREADY_EXISTS` instead of becoming the reparse write target.
- The create-race path also returns `ERROR_ALREADY_EXISTS` when an existing directory appears without a grant.
- `MovePathInternal(...)` now computes `directoryMergeDestination` only for non-reparse source directories, preventing source junctions from entering `MoveDirectoryMergeByRename(...)`.
- The old Clearflow reparse test no longer encodes unsafe no-grant conversion; it proves explicit overwrite still creates the destination junction.
