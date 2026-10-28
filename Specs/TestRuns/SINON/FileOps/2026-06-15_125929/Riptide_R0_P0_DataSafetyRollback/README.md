# Riptide R0 P0 Data-Safety / Rollback Evidence

Date: 2026-06-15
Machine: SINON
Branch: `claude/zealous-poitras-ce383a`

## Summary

- Build after R0-1/R0-2 code changes: `.\build.ps1 -ProjectName RedSalamander` passed with 0 warnings and 0 errors.
- R0-1 RED: `Riptide_MoveDistinctSameSizeFilePreservesSource` failed on the trunk-integrated gate commit with only the new selftests applied. Failure: old engine completed `S_OK`, showed 0 prompts, and deleted/changed the source.
- R0-1 GREEN: exact case passed on current engine after replacing size+mtime identity with no-follow byte comparison and same-file identity validation.
- R0-2 RED: `Riptide_ReparseMoveRollbackKeepsOverwrittenDestination` failed before the rollback fix. Failure: pre-existing overwritten destination was deleted after injected source-delete failure.
- R0-2 GREEN: exact case passed after `ShouldRollbackCopiedDestination(destinationExistedBeforeCopy, sourceContentDeleted)`.
- Neighbor checks passed: `Fairstream_RerunCopyIsResumeAware`, `Fairstream_MoveFallbackPreservesUncopiedNewFiles`, `Fairstream_ReparseReplaceNonEmptyDirRequiresConsent`, `Fairstream_MovedTreeRetargetsInternalLinks`, `FileOps_ReparseRetargetDestinationContainment`, and `FileOps_CrossVolumeMovePartialFailureStatus`.
- Broad FileOps immediately after P0 was red: 60 passed, 3 failed, 39 skipped. `Phase6_LocalBandwidthThrottle` and `Phase7_LargeDirectoryEnumeration` were then fixed and pass in focused reruns. Remaining known broad-suite blocker is tracked in the WIP plan as R3-5.

## Local Artifact Folders

The raw `last_run` folders are intentionally ignored by Git, but were copied here locally for same-machine review:

- `R0-1_RED_selftests_only_old_engine/`
- `R0-1_GREEN_current_engine/`
- `R0-2_RED_current_engine_before_fix/`
- `R0-2_GREEN_current_engine/`
- `P0_neighbor_checks/`
- `P0_broad_FileOps_red/`
- `GraphBands_exact_red/`
- `Phase6_LocalBandwidthThrottle_isolated_red/`
- `Phase6_LocalBandwidthThrottle_short_root_timeout_red/`
- `Phase6_LocalBandwidthThrottle_normal_root_green/`
- `Phase6_LocalBandwidthThrottle_short_root_green/`
- `Phase7_LargeDirectoryEnumeration_long_root_red/`
- `Phase7_LargeDirectoryEnumeration_normal_root_green/`
- `Phase7_LargeDirectoryEnumeration_short_root_green/`
