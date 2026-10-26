# BR-3 Fresh Full evidence

Source: c36d3497a0d5214168c062bf3f67a16e6baa909a plus the P1/P2/BR-3 worktree.
Command: Tools/Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -AllowAlternateVolumeTestRoot -TestRoot Z:\RedSalamander.Perf
Runner: 20260908T073251Z-67752-50d9f956c380407d9fb1ef91d0b66c49.
Profile: test-enabled x64 Debug; every build/re-attestation reported zero warnings/errors.
Accepted final build receipt: d9c4193d03465a6dae3cbf594d748a63b797012eb44e0bcee511e3d39517ba1e.
Workspace snapshot: 9f0b29ed1d231ec496e9f15469bda93ffc276b96992024866339172143ad3dda.

## Verdict and scope

This Fresh Full invocation executed all 35 entries: 34 were promoted and Commands
failed. Repository and change verdicts are FAILED, with 2,108 passes, one failure
and 52 explicit skips in the runner counts. Duration: 3,924,375 ms. There was no
source-mutation invalidation and no reused or provisional entry. This is not a
green repository-wide closeout or composite/Resume evidence.

The FileOperations entry passed in the complete process sequence: 178 passed,
zero failed, 20 skipped (18 missing remote-profile checks; two 7z cases owned by
Compare). Every case and skip reason is retained. The BR-3 witnesses passed:

| Case | Duration |
|---|---:|
| BR3_CancelQueuePausedTransfer | 9,313 ms fixture; 78,000 us cancellation with predecessor still paused |
| Phase7_ParallelCopyMoveKnobs | 38,594 ms |
| Phase7_SharedPerItemScheduler | 1,047 ms |
| C1_ExitCloseDeferredUntilTasksQuiet | 875 ms |
| Phase14_PopupHostLifetimeGuard | 391 ms |
| R4A19_DiscoveryIndependentVolumes | 3,750 ms; configured Z: and D: roots |

ConflictCancelWake has value0=1/value1=0. CopyMoveKnobObserved proves three
in-flight files with three active calls at configured concurrency 4, and two
in-flight files with eight active calls at configured concurrency 8; both rows
have value1=1 for observed per-call sharing. The fixture releases the speed limit
only after those observations. No historical worker-stack location or throughput
improvement is inferred from these controls.

## Remaining Full failure

Commands: cmd_preferences_dialog_keyboard_reordered_resized_copy_follows_visible_columns_after_search_roundtrip.
Reason: Preferences Keyboard filtered no-match search rebuild did not settle
during reordered-resized-copy/search validation. Duration 3,690 ms; runner
classification UNCLASSIFIED_FAILURE, no retry classification evidence. Commands
otherwise recorded 893 passes and two opt-in skips. Compare recorded 242 passes,
zero failures and 30 prerequisite skips. Their complete case results are included.
This Preferences case is outside the modified File Operations code and remains
an open Full-gate issue; passing FileOperations must not conceal it.

## Retained artifacts

The aggregate v2 and v1 sidecar, immutable plan, terminal run state, workspace
snapshot and case history preserve verdicts, plan/decision identities, per-entry
fingerprints, promotion digests, accepted receipt and all declared result reasons.
Results and traces are copied verbatim. The performance selection retains all
FileOps.SelfTest rows and the specified Phase 7 scheduler/throttle/interlock
windows in original order. perf-compaction.json binds the original byte/row
counts and SHA-256 plus the exact selection and output digest. No distribution
or throughput claim is derived from this bounded selection.