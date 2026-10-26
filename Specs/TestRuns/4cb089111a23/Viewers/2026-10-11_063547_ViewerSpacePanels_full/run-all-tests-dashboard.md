# Run-AllTests Case Dashboard

- Run: `20261011T040752Z-65208-c8d6028187b24b18b10e1ff93622171b`
- Suite: `Full`

## Top Slowest Cases

| Harness | Case | Source | Status | Classification | Duration ms | Seed | Attempt |
|---|---|---|---|---|---:|---|---:|
| Commands | cmd_viewer_space_hover_does_not_rebuild_static_cache | case | passed | UNCLASSIFIED_FAILURE | 74492 |  | 1 |
| Commands | cmd_compare_directories_options_long_run_open_close_stays_stable | case | passed | UNCLASSIFIED_FAILURE | 56250 |  | 1 |
| Commands | cmd_pane_rename_folder_prompt_batch_button_opens_batch_rename | case | passed | UNCLASSIFIED_FAILURE | 30127 |  | 1 |
| Commands | cmd_preferences_dialog_category_switches_do_not_churn_tree_host | case | passed | UNCLASSIFIED_FAILURE | 29298 |  | 1 |
| Commands | cmd_pane_fileops_popup_resource_lifetime | case | passed | UNCLASSIFIED_FAILURE | 25481 |  | 1 |
| Commands | folderView_column_widths_audit | case | passed | UNCLASSIFIED_FAILURE | 24043 |  | 1 |
| Commands | folderView_perf_overlay_invalidation_stress | case | passed | UNCLASSIFIED_FAILURE | 20461 |  | 1 |
| Commands | settings_save_queue_serializes_coalesces_and_flushes | case | passed | UNCLASSIFIED_FAILURE | 20289 |  | 1 |
| Commands | cmd_compare_directories_options_live_dx_body_interaction | case | passed | UNCLASSIFIED_FAILURE | 17928 |  | 1 |
| Commands | dispatch_smoke_all_commands | case | passed | UNCLASSIFIED_FAILURE | 14916 |  | 1 |
| Commands | cmd_viewer_space_memory_settles_after_close | case | passed | UNCLASSIFIED_FAILURE | 13723 |  | 1 |
| Commands | cmd_pane_fileops_popup_throughput_restart | case | passed | UNCLASSIFIED_FAILURE | 12511 |  | 1 |
| Commands | cmd_viewer_space_explorer_analyst_exact_nested_detail | case | failed | UNCLASSIFIED_FAILURE | 11279 |  | 1 |
| Commands | cmd_pane_fileops_popup_throughput_disclosure | case | failed | UNCLASSIFIED_FAILURE | 10963 |  | 1 |
| Commands | cmd_pane_batchRename_admission_worker_owned_cancellable | case | passed | UNCLASSIFIED_FAILURE | 10377 |  | 1 |
| Commands | cmd_pane_find_dialog_result_shortcuts_use_shell_clipboard_and_file_actions | case | failed | UNCLASSIFIED_FAILURE | 10212 |  | 1 |
| Commands | cmd_connection_manager_window_tab_traversal_live_dx_interaction | case | passed | UNCLASSIFIED_FAILURE | 9577 |  | 1 |
| Commands | cmd_preferences_dialog_shell_footer_live_dx_interaction | case | failed | UNCLASSIFIED_FAILURE | 9209 |  | 1 |
| Commands | cmd_preferences_dialog_themes_roundtrip_restores_dxui_surface | case | passed | UNCLASSIFIED_FAILURE | 9207 |  | 1 |
| Commands | cmd_preferences_dialog_category_tree_handles_reverse_keyboard_navigation | case | passed | UNCLASSIFIED_FAILURE | 8586 |  | 1 |

## Failure And Retry Evidence

| Harness | Case | Source | Status | Classification | Seed | Reason |
|---|---|---|---|---|---|---|
| ToolsPesterTests | keeps every approved post-Terminal Done history blob unchanged | case | failed | UNCLASSIFIED_FAILURE |  |  |
| CompareDirectories | remote_file_s3 | case | skipped | PASSED |  | S3: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_S3 or create 'FileOpsSelfTest S3'). |
| CompareDirectories | remote_s3_directory_size_callback_contract | case | skipped | PASSED |  | S3: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_S3 or create 'FileOpsSelfTest S3'). |
| CompareDirectories | remote_s3_pagination | case | skipped | PASSED |  | S3: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_S3 or create 'FileOpsSelfTest S3'). |
| CompareDirectories | remote_file_onedrive_personal | case | skipped | PASSED |  | OneDrive Personal: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_ONEDRIVE_PERSONAL or create 'FileOpsSelfTest OneDrive Personal'). |
| CompareDirectories | remote_onedrive_personal_directory_size_callback_contract | case | skipped | PASSED |  | OneDrive Personal: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_ONEDRIVE_PERSONAL or create 'FileOpsSelfTest OneDrive Personal'). |
| CompareDirectories | remote_file_onedrive_business | case | skipped | PASSED |  | OneDrive Business: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_ONEDRIVE_BUSINESS or create 'FileOpsSelfTest OneDrive Business'). |
| CompareDirectories | remote_onedrive_business_directory_size_callback_contract | case | skipped | PASSED |  | OneDrive Business: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_ONEDRIVE_BUSINESS or create 'FileOpsSelfTest OneDrive Business'). |
| CompareDirectories | remote_file_sharepoint | case | skipped | PASSED |  | SharePoint: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_SHAREPOINT or create 'FileOpsSelfTest SharePoint'). |
| CompareDirectories | remote_sharepoint_directory_size_callback_contract | case | skipped | PASSED |  | SharePoint: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_SHAREPOINT or create 'FileOpsSelfTest SharePoint'). |
| CompareDirectories | remote_file_ftp | case | skipped | PASSED |  | FTP: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_FTP or create 'FileOpsSelfTest FTP'). |
| CompareDirectories | remote_file_sftp | case | skipped | PASSED |  | SFTP: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_SFTP or create 'FileOpsSelfTest SFTP'). |
| CompareDirectories | remote_file_scp | case | skipped | PASSED |  | SCP: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_SCP or create 'FileOpsSelfTest SCP'). |
| CompareDirectories | remote_file_imap | case | skipped | PASSED |  | IMAP: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_IMAP or create 'FileOpsSelfTest IMAP'). |
| CompareDirectories | remote_file_smb | case | skipped | PASSED |  | SMB smoke skipped: configure path.smb.primary or set REDSALAMANDER_SELFTEST_SMB_ROOT. |
| CompareDirectories | remote_ftp_directory_size_callback_contract | case | skipped | PASSED |  | FTP: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_FTP or create 'FileOpsSelfTest FTP'). |
| CompareDirectories | remote_ftp_continue_on_error_partial | case | skipped | PASSED |  | FTP: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_FTP or create 'FileOpsSelfTest FTP'). |
| CompareDirectories | remote_s3_metadata_smoke | case | skipped | PASSED |  | S3: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_S3 or create 'FileOpsSelfTest S3'). |
| CompareDirectories | remote_s3_delete_missing | case | skipped | PASSED |  | S3: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_S3 or create 'FileOpsSelfTest S3'). |
| CompareDirectories | mtp_live_device_smoke | case | skipped | PASSED |  | MTP live smoke skipped: set REDSALAMANDER_SELFTEST_MTP_DEVICE to run against an approved device. |
| CompareDirectories | local_index_core_sqlite_option_keeps_snapshot_runtime_store | case | skipped | PASSED |  | Live journal cursor unavailable; direct SQLite query cutover cannot prove currentness. |
| CompareDirectories | local_index_core_sqlite_cold_start_bypasses_snapshot_runtime_store | case | skipped | PASSED |  | Live journal cursor unavailable; direct SQLite query cutover cannot prove currentness. |
| CompareDirectories | local_index_core_sqlite_authoritative_replays_without_snapshot_runtime_store | case | skipped | PASSED |  | Live journal cursor unavailable; direct SQLite query cutover cannot prove currentness. |
| CompareDirectories | local_index_core_sqlite_authoritative_reseeds_after_store_loss_during_replay | case | skipped | PASSED |  | Live journal cursor unavailable; direct SQLite query cutover cannot prove currentness. |
| CompareDirectories | local_index_core_sqlite_unrelated_pending_legacy_keeps_requested_ready | case | skipped | PASSED |  | Live journal cursor unavailable; direct SQLite query cutover cannot prove currentness. |
| CompareDirectories | local_index_core_sqlite_prefilter_classifies_name_patterns | case | skipped | PASSED |  | Live journal cursor unavailable; direct SQLite query cutover cannot prove currentness. |
| CompareDirectories | search_service_sqlite_ntfs_traversal_seed_stays_degraded | case | skipped | PASSED |  | NTFS traversal-seed degradation test requires an NTFS-backed self-test root. |
| CompareDirectories | search_service_candidate_impersonation_failure_is_incomplete_warning | case | skipped | PASSED |  | Search service test hooks are unavailable in this build configuration. |
| CompareDirectories | host_fallback_search_remote_ftp_name_only | case | skipped | PASSED |  | FTP: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_FTP or create 'FileOpsSelfTest FTP'). |
| CompareDirectories | local_search_service_root_rejection_does_not_arm_transport_cooldown | case | skipped | PASSED |  | Search service test hooks are unavailable in this build configuration. |
| CompareDirectories | search_service_transient_parent_failure_is_cached_and_warns_batch_and_completion | case | skipped | PASSED |  | Search service test hooks are unavailable in this build configuration. |
| Commands | file_operations_phase0_typed_contract_source_guard | case | failed | UNCLASSIFIED_FAILURE |  | Clipboard Move must be consumed at accepted admission, never deferred to terminal success. |
| Commands | cmd_preferences_dialog_shell_footer_live_dx_interaction | case | failed | UNCLASSIFIED_FAILURE |  | Preferences General page did not refocus the menu-bar toggle before restoration. |
| Commands | cmd_pane_find_dialog_result_shortcuts_use_shell_clipboard_and_file_actions | case | failed | UNCLASSIFIED_FAILURE |  | Failed to quiesce pre-existing file operations before Find shortcut file-action validation. |
| Commands | cmd_compare_directories_options_uses_dxui_labels_without_visible_legacy_statics | case | failed | UNCLASSIFIED_FAILURE |  | Compare Directories options should render exactly one visible DxUi body host during the initial baseline DX surface probe. |
| Commands | cmd_pane_fileops_popup_throughput_disclosure | case | failed | UNCLASSIFIED_FAILURE |  | Focused path must reveal its exact labelled value without activating a task decision: id=FileOperations.Task.90001.Path.Source.0 actual=FileOperations.Task.90001.Path.Source.0 pathUnits=128 tooltipUnits=0 expectedUnits=232 tooltipPrefix=''. |
| Commands | cmd_pane_fileops_conflict_two_pending_task_identity | case | failed | UNCLASSIFIED_FAILURE |  | The task disclosure must expose a measured pointer target. |
| Commands | cmd_viewer_space_explorer_analyst_exact_nested_detail | case | failed | UNCLASSIFIED_FAILURE |  | Owned-window capture explorer-light retains chrome, analysis and navigation text: 0x8007000D; analysisPixels=1802, navigationPixels=336. |
| Commands | cmd_viewer_space_cache_reopen_200_first_map_optin | case | skipped | UNCLASSIFIED_FAILURE |  | Set REDSALAMANDER_VIEWERSPACE_CACHED_REOPEN_200=1 for the opt-in 200-sample cached first-map benchmark. |
| Commands | cmd_viewer_space_memory_heavy_four_root_optin | case | skipped | UNCLASSIFIED_FAILURE |  | Set REDSALAMANDER_VIEWERSPACE_MEMORY_HEAVY=1 for the opt-in four-root 100k memory benchmark. |
| Commands | cmd_viewer_space_memory_heavy_four_root_retained_reference_optin | case | skipped | UNCLASSIFIED_FAILURE |  | Set REDSALAMANDER_VIEWERSPACE_MEMORY_HEAVY=1 for the opt-in four-root 100k memory benchmark. |
| Commands | cmd_viewer_space_layout_20k_visible_optin | case | skipped | UNCLASSIFIED_FAILURE |  | Set REDSALAMANDER_VIEWERSPACE_20K_LAYOUT=1 to run the opt-in ViewerSpace 20k layout scenario. |
| Commands | cmd_viewer_space_layout_50k_visible_optin | case | skipped | UNCLASSIFIED_FAILURE |  | Set REDSALAMANDER_VIEWERSPACE_50K_LAYOUT=1 to run the opt-in ViewerSpace 50k visible-layout scenario. |
| Commands | viewer_space_perf_large_optin | case | skipped | UNCLASSIFIED_FAILURE |  | Set REDSALAMANDER_VIEWERSPACE_LARGE_PERF=1 to run the opt-in ViewerSpace large perf scenario. |
| Commands | folderView_thumbnail_sort_popup_slider | case | failed | UNCLASSIFIED_FAILURE |  | Thumbnail size slider drag should update the pane through the real size command. |
| FileOperations | FileOps_QueuedAdmissionContracts | case | failed | UNCLASSIFIED_FAILURE |  | A change refused for a destination name must keep the destination and name the blocking name on the card and in its diagnostic (hr=0x8007007B, name='trailing-dot.', diagnostic=false). |
| FileOperations | FileOps_SafetyRemediationContracts | case | failed | UNCLASSIFIED_FAILURE |  | File Operations safety contracts failed; see individual witnesses in the trace. |
| FileOperations | FileOps_AdmissionBaseline | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_OwnedProductionShutdown | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_CrossVolumeMoveAfterPartialPublicationStorage | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_RepeatedNonClipboardCopyStorage | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_RepeatedNonClipboardMoveStorage | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_TerminalFactsProjection | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_ClearCompletedScope | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_PerItemCancelUndispatchedRoots | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_AdmissionFailureReachesPane | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_DeferredExitRefusesNewWork | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_CompareWindowCloseConsent | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_PopupFailureSurfaceRecovery | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_PopupStableHostedTree | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_PopupIdleCadence | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_AlertOverlayLifetime | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_AlertOverlayUxAccessibility | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_BatchRenameExactResultCarrier | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_BatchRenameNeverRunRowsNotAttempted | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_CopyMergeIntoExistingFolder | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_MoveMergeIntoExistingFolderSameVolume | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Beeline_RenameMergeSkipKeepsSourceFolder | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_RenameMergeEmptyFolderCompletes | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_ReparseDirectoryMergeIntoExistingFolder | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_ProviderCapabilityMatrix | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_ParallelGraphFairColorWeight | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_CopyOnlyMoveRetainsSourceWithoutDestinationProof | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_ReparseLiteralCopyKeepsTargets | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_DeleteToctouSwapGuard | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_NativeMoveRacedDestinationNotReplaced | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_NativeMoveMissingPathKnownFailure | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | FileOps_ResolvedItemsExactDestinations | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Riptide_MoveDistinctSameSizeFilePreservesSource | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Riptide_ReparseNativeMoveRelocatesLinkObject | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Fairstream_ManagedMovePreservesUncopiedNewFiles | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Fairstream_MoveMergeReadonlyDestinationFolder | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Fairstream_OverwriteGrantIsOneShot | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Fairstream_ConflictPromptsSerializedUnderParallelism | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Fairstream_ReparseReplaceNonEmptyDirRequiresConsent | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Floodgate_ConflictApplyAllKeepsRiskBucketsSeparate | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Riptide_ReparseCopyOntoEmptyRealDirRequiresConsent | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Riptide_ReparseMoveSourceNeverUsesDirectoryMerge | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Fairstream_MovedTreeKeepsLiteralLinks | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Fairstream_CrossFsConcurrentMoveUsesBridge | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Fairstream_RerunCopyIdenticalCollisionPrompts | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Fairstream_MoveSameSizeCollisionPrompts | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Fairstream_GraphBandsFairUnderParallelCopy | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Fairstream_ManagedMergeMovePreservesChildren | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Fairstream_JunctionMergeTargetRequiresConsent | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Fairstream_CopyIntoSelfAliasRejected | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Fairstream_StorageKindProbed | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Fairstream_ParallelDeleteContinuesPastLockedChild | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Fairstream_BridgePerFileConflictSkips | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Cinderstar_BridgeMovePerFileConflictSkipsParallel | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Riptide_BridgeNestedDirVsFileSkipContinuesSiblings | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Riptide_BridgeDirectoryOverReadOnlyFileIsTypeMismatch | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Riptide_BridgeCreateDirectoryRaceExistingFilePromptsPartial | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Fairstream_SaturationConcurrentCopiesMakeProgress | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Fairstream_DiscoveryAheadOverlapsTransfer | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Riptide_LiveFinishedSnapshotCarriesDiagnostics | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Riptide_BridgeSequentialContinueOnErrorCopiesSiblings | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Causeway_BridgeRejectsHostileChildNames | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Causeway_BridgeProviderOutputContracts | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Causeway_BridgeSchedulingAndResourceContracts | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Causeway_BridgeFileReparsePolicy | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Beeline_CopySkipLinksKeepsPlaceholders | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Causeway_BridgeFailureStatusAndPausedReader | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Floodgate_CrossFsCopyGetSizeFailureRefusesCommit | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Floodgate_CrossFsMoveGetSizeFailurePreservesSource | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Floodgate_CrossFsMoveDestinationGetSizeFailurePreservesSource | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Cinderstar_LegacyWriterEventuallyConsistentMove | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Cinderstar_LegacyWriterPermanentMissMove | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Cinderstar_LegacyWriterWrongSizeMove | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Cinderstar_LegacyWriterConcurrentReplacementCopy | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Cinderstar_LegacyWriterCancelDuringBackoff | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Floodgate_CrossFsCopyOnlyUsesCopyVerification | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Floodgate_CrossFsMoveGetSizeFailuresParallelPreserveSourceTree | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Floodgate_CrossFsCopyOnlyNeverEntersSourceCleanup | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Floodgate_CrossFsDirectoryCopyOnlyRetainsSource | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Floodgate_LocalWriterOverwriteIsStaged | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Floodgate_LocalCopyOverwriteIsStaged | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Floodgate_LocalCopyNewNameConcurrentReplacementSurvives | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Floodgate_LocalCopyNewNameAbortFailureIsRetainedIncomplete | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Riptide_HostPerItemSchedulerShutdownWaitsForBlockedWorker | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Floodgate_InlineF2WorkerQueueBypassAndInterlock | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | C1_InlineRenameNameRefusedOnCard | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | R4A19_DiscoveryMeasurementFacts | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | R4A19_DiscoveryProviderControls | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | R4A19_DiscoveryIndependentVolumes | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | DiscoveryScope_LeafCopy | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | DiscoveryScope_LeafMove | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | DiscoveryScope_ManagedDirectory | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | DiscoveryScope_NativeRenameMerge | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | DiscoveryScope_MixedRoots | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | DiscoveryScope_ZeroByteAndVerifiedLeaf | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | DiscoveryScope_LocalDirectApi | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | DiscoveryScope_DummyDirectApi | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | DiscoveryScope_CurlDirectApi | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | DiscoveryScope_S3DirectApi | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | DiscoveryScope_GDriveDirectApi | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | DiscoveryScope_GraphDirectApi | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | DiscoveryScope_MtpDirectApi | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | R0fSmb_BlockedSynchronousCallCancelReturns | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | R0fSmb_LoopbackReadWriteCreateDelete | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | R0fSmb_LoopbackRecycleNeverDestroys | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | R0fCurl_FakeFtpReadWriteCreateDelete | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | BR5_CurlHostDirectoryMove | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | BR5_CurlHostDirectoryMoveRefused | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | R0fS3_FakeS3ReadWriteCreateDelete | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | R0fGraph_FakeGraphReadWriteCreateRenameRecycle | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | R0fGDrive_FakeDriveReadWriteCreateMoveDelete | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | C0_GDriveCommittedCopyResponseLost | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | C0_GDriveCommittedDeleteResponseLost | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | C0_NativeCopyFailurePreservesKnownAxes | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | C0_MutationReceiptPrefixBoundary | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | C0_LocalKnownNoCommitReceipts | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | C0_CurlNativeDeleteGuards | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | C0_CurlCommittedMutationResponseLost | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | C0_CurlHostCommittedDeleteResponseLost | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | C0_CurlPartialTreeFailure | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | C0_CurlHostPartialTreeFailure | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | C0_CurlNativeDeleteLateListing | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | C0_CurlDeleteTraversalBounds | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | C0_CurlDirectorySizeTruth | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | C0_CurlMovePreflightTruth | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | C0_CurlCopyTraversalTruth | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | C0_CurlEntryLookupTruth | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | C0_CurlImapListingTruth | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | C0_CurlImapTransportTruth | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | R3_1_IdentityLessReplaceDummy | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | R3_2_WriterProofDummy | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | R3_3_PublicationFaultMatrix | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | R3_4_ReplaceWithoutOccupantTokenRefused | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | RC3_9_TransferDestinationNameRefused | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | R4T_DeepTreeCopyCompletesIteratively | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | R4T2_DeepTreeLocalDeleteCompletes | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | R4T3_WideDirectoryCopyCompletes | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | R4A02_DeleteOverCopySourceWarnsAndQueues | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | R4A02_DontStartCancelsBeforeMutation | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | R4A02_ReadReadDoesNotWarn | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | R4A02_SameDestinationWarns | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | R4A02_RunConcurrentSameDestinationStaysSafe | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | R4A02_LiveOutputGuardChoices | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | R4A02_DisjointWritesDoNotWarn | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | R4A02_RenamePublishKeepsIndexPrecise | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | R4A02_LatePublisherStillWarns | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase5_DiscoverySingleTraversal | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase5_DiscoveryCancelReleasesSlot | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase5_CleanupDrainsOutstandingTasks | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase5_DiscoveryCancelLatencyLocal | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase5_DiscoverySkipContinues | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | BR3_CancelQueuePausedTransfer | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase5_CancelQueuedTask | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase5_SwitchParallelToWaitDuringDiscovery | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase5_SwitchWaitToParallelResume | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase6_PopupRateSmoothing | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase6_PopupSmokeResizeAndPause | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase6_DeleteBytesMeaningful | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase6_LocalBandwidthThrottle | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase6_ParallelBandwidthThrottleFairness | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase7_WatcherChurn | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase7_CacheBorrowNoWatchInvalidation | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase7_CrossPaneVisibleRefreshLocal | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase7_CrossPaneVisibleRefreshDummy | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase7_CrossPaneRelocateLocal | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase7_LargeDirectoryEnumeration | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase7_ParallelCopyMoveKnobs | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase7_CopyMoveConcurrency16Perf | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase7_AutoConcurrencyHints | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase7_PerItemDirectoryCopyInFlightLines | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase7_CopyItemsSingleFolderRecursiveParallelism | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase7_CopyItemsMultiRootUnevenRecursiveParallelism | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase7_CopyRecursiveParallelismMatrix | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase7_SharedPerItemScheduler | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase7_ParallelDeleteKnobs | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase7_RecycleBinBatchDelete | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase7_RecycleBinBatchDeleteMultiBatch | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase7_RecycleBinTestItemsPurged | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase8_DefaultBandwidthLimitFromSettings | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase8_TightDefaults_NoOverwrite | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase8_InvalidDestinationRejected | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase8_InvalidSizeBytesRejected | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase8_PerItemOrchestration | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase9_ConflictPrompt_OverwriteReplaceReadonly | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase9_ConflictPrompt_ApplyToAllUiCache | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase9_ConflictPrompt_KeepBothNestedCacheEligibility | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase9_ConflictPrompt_TypeMismatchNoOverwrite | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase9_ConflictPrompt_LocalFileOntoDirectory | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase9_ConflictPrompt_SkipApplyToAll | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase9_ConflictPrompt_RetryCap | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase9_ConflictPrompt_SkipContinuesDirectoryCopy | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase9_PerItemConcurrency | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase10_PermanentDelete | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase10_DeferredConsentAndRecycleEscalation | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase10_RecycleDestroyedIsNeverRecycled | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase10_RecycleEscalationDirectoryTree | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase10_RecycleEscalationTaskScopedGrant | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase10_RecycleEscalationBulkItemsContinueOnError | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase10_RecycleEscalationBulkItemsResumeBounded | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase10_MetadataPreservationAndSourceRetention | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase10_TypedResultsAndConsumers | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase10_ClipboardAdmissionAndRetainedActions | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase10_ContentVerification | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase10_ArtifactTouchGuard | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | R1d_PreparingLifecycle | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | C1_ExitCloseDeferredUntilTasksQuiet | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Beeline_SameVolumeTreeMoveIsRename | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Beeline_LoopbackShareMoveIsRename | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | C1_PermanentDeleteConfirmsOnCard | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase11_CrossFileSystemBridge | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase11_BridgeSingleFolderParallelCopyInFlightLines | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase11_BridgeMultiFolderParallelCopyInFlightLines | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase11_BridgePipelineDummyToDummyPerf | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase11_ConnectionOverridePrecedence | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase11_ConnectionOverrideGlobalGate | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase11_ConnectionOverrideClamp | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase12_ReparsePointPolicy | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase13_PostMortemDiagnostics | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase14_PopupHostLifetimeGuard | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase15_FileSystem7zReadSeekSmoke | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase15_FileSystem7zMountPathImpact | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase16_RemoteWatchContractExposure | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase16_RemoteFtpSecret | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase16_RemoteFtpSandbox | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase16_RemoteSftpSecret | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase16_RemoteSftpSandbox | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase16_RemoteScpSecret | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase16_RemoteScpSandbox | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase16_RemoteImapSecret | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase16_RemoteImapSandbox | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase16_RemoteS3Secret | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase16_RemoteS3Sandbox | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase16_RemoteS3FileOps | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase16_RemoteOneDrivePersonalSecret | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase16_RemoteOneDrivePersonalSandbox | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase16_RemoteOneDrivePersonalFileOps | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase16_RemoteOneDriveBusinessSecret | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase16_RemoteOneDriveBusinessSandbox | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase16_RemoteSharePointSecret | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
| FileOperations | Phase16_RemoteSharePointSandbox | case | skipped | UNCLASSIFIED_FAILURE |  | not executed (suite stopped early) |
