# Run-AllTests Case Dashboard

- Run: `20260908T200723Z-69288-1db8a2860de7474ebbc2f9164aa99216`
- Suite: `Full`

## Top Slowest Cases

| Harness | Case | Source | Status | Classification | Duration ms | Seed | Attempt |
|---|---|---|---|---|---:|---|---:|
| FileOperations | Phase11_BridgeMultiFolderParallelCopyInFlightLines | case | passed | PASSED | 196078 |  | 1 |
| FileOperations | C0_CurlCopyTraversalTruth | case | passed | PASSED | 182891 |  | 1 |
| FileOperations | Phase11_BridgePipelineDummyToDummyPerf | case | passed | PASSED | 128687 |  | 1 |
| FileOperations | R4T3_WideDirectoryCopyCompletes | case | passed | PASSED | 120781 |  | 1 |
| Commands | cmd_pane_batchRename_admission_worker_owned_cancellable | case | passed | PASSED | 89212 |  | 1 |
| Commands | cmd_pane_batchRename_engine_large_preview_perf | case | passed | PASSED | 53563 |  | 1 |
| Commands | cmd_compare_directories_options_long_run_open_close_stays_stable | case | passed | PASSED | 53375 |  | 1 |
| FileOperations | R4T_DeepTreeCopyCompletesIteratively | case | passed | PASSED | 47703 |  | 1 |
| FileOperations | Phase11_ConnectionOverrideGlobalGate | case | passed | PASSED | 40422 |  | 1 |
| FileOperations | Phase7_ParallelCopyMoveKnobs | case | passed | PASSED | 38172 |  | 1 |
| Commands | folderView_perf_focus_selection_state | case | passed | PASSED | 36815 |  | 1 |
| FileOperations | Phase6_PopupSmokeResizeAndPause | case | passed | PASSED | 35922 |  | 1 |
| CompareDirectories | local_search_service_disconnect_falls_back_local_index | case | passed | PASSED | 33848 |  | 1 |
| Commands | folderView_perf_overlay_invalidation_stress | case | passed | PASSED | 33408 |  | 1 |
| Commands | cmd_preferences_dialog_category_switches_do_not_churn_tree_host | case | passed | PASSED | 31305 |  | 1 |
| Commands | cmd_pane_rename_folder_prompt_batch_button_opens_batch_rename | case | passed | PASSED | 30222 |  | 1 |
| FileOperations | C0_CurlEntryLookupTruth | case | passed | PASSED | 26547 |  | 1 |
| Commands | folderView_column_widths_audit | case | passed | PASSED | 24951 |  | 1 |
| FileOperations | Phase11_BridgeSingleFolderParallelCopyInFlightLines | case | passed | PASSED | 24937 |  | 1 |
| FileOperations | C0_CurlMovePreflightTruth | case | passed | PASSED | 20797 |  | 1 |

## Failure And Retry Evidence

| Harness | Case | Source | Status | Classification | Seed | Reason |
|---|---|---|---|---|---|---|
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
| CompareDirectories | local_index_core_sqlite_cutover_blocked_by_pending_legacy_import | case | skipped | PASSED |  | Live journal cursor unavailable; direct SQLite query cutover cannot prove currentness. |
| CompareDirectories | local_index_core_sqlite_prefilter_classifies_name_patterns | case | skipped | PASSED |  | Live journal cursor unavailable; direct SQLite query cutover cannot prove currentness. |
| CompareDirectories | search_service_sqlite_ntfs_traversal_seed_stays_degraded | case | skipped | PASSED |  | NTFS traversal-seed degradation test requires an NTFS-backed self-test root. |
| CompareDirectories | search_service_sqlite_query_failure_falls_back_live_scan | case | skipped | PASSED |  | Live journal cursor unavailable; direct SQLite query cutover cannot prove currentness. |
| CompareDirectories | search_service_sqlite_midquery_failure_restarts_live_scan_without_duplicates | case | skipped | PASSED |  | Live journal cursor unavailable; direct SQLite query cutover cannot prove currentness. |
| CompareDirectories | search_service_sqlite_prefilter_roundtrip | case | skipped | PASSED |  | Live journal cursor unavailable; direct SQLite query cutover cannot prove currentness. |
| CompareDirectories | host_fallback_search_remote_ftp_name_only | case | skipped | PASSED |  | FTP: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_FTP or create 'FileOpsSelfTest FTP'). |
| Commands | cmd_viewer_space_layout_20k_visible_optin | case | skipped | PASSED |  | Set REDSALAMANDER_VIEWERSPACE_20K_LAYOUT=1 to run the opt-in ViewerSpace 20k layout scenario. |
| Commands | viewer_space_perf_large_optin | case | skipped | PASSED |  | Set REDSALAMANDER_VIEWERSPACE_LARGE_PERF=1 to run the opt-in ViewerSpace large perf scenario. |
| FileOperations | Phase15_FileSystem7zReadSeekSmoke | case | skipped | PASSED |  | 7z read/seek smoke skipped in the file-operations suite; 7z coverage remains in compare selftests. |
| FileOperations | Phase15_FileSystem7zMountPathImpact | case | skipped | PASSED |  | 7z mount-path impact skipped in the file-operations suite; 7z coverage remains in compare selftests. |
| FileOperations | Phase16_RemoteFtpSecret | case | skipped | PASSED |  | FTP: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_FTP or create 'FileOpsSelfTest FTP'). |
| FileOperations | Phase16_RemoteFtpSandbox | case | skipped | PASSED |  | FTP: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_FTP or create 'FileOpsSelfTest FTP'). |
| FileOperations | Phase16_RemoteSftpSecret | case | skipped | PASSED |  | SFTP: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_SFTP or create 'FileOpsSelfTest SFTP'). |
| FileOperations | Phase16_RemoteSftpSandbox | case | skipped | PASSED |  | SFTP: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_SFTP or create 'FileOpsSelfTest SFTP'). |
| FileOperations | Phase16_RemoteScpSecret | case | skipped | PASSED |  | SCP: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_SCP or create 'FileOpsSelfTest SCP'). |
| FileOperations | Phase16_RemoteScpSandbox | case | skipped | PASSED |  | SCP: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_SCP or create 'FileOpsSelfTest SCP'). |
| FileOperations | Phase16_RemoteImapSecret | case | skipped | PASSED |  | IMAP: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_IMAP or create 'FileOpsSelfTest IMAP'). |
| FileOperations | Phase16_RemoteImapSandbox | case | skipped | PASSED |  | IMAP: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_IMAP or create 'FileOpsSelfTest IMAP'). |
| FileOperations | Phase16_RemoteS3Secret | case | skipped | PASSED |  | S3: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_S3 or create 'FileOpsSelfTest S3'). |
| FileOperations | Phase16_RemoteS3Sandbox | case | skipped | PASSED |  | S3: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_S3 or create 'FileOpsSelfTest S3'). |
| FileOperations | Phase16_RemoteS3FileOps | case | skipped | PASSED |  | S3: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_S3 or create 'FileOpsSelfTest S3'). |
| FileOperations | Phase16_RemoteOneDrivePersonalSecret | case | skipped | PASSED |  | OneDrive Personal: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_ONEDRIVE_PERSONAL or create 'FileOpsSelfTest OneDrive Personal'). |
| FileOperations | Phase16_RemoteOneDrivePersonalSandbox | case | skipped | PASSED |  | OneDrive Personal: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_ONEDRIVE_PERSONAL or create 'FileOpsSelfTest OneDrive Personal'). |
| FileOperations | Phase16_RemoteOneDrivePersonalFileOps | case | skipped | PASSED |  | OneDrive Personal: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_ONEDRIVE_PERSONAL or create 'FileOpsSelfTest OneDrive Personal'). |
| FileOperations | Phase16_RemoteOneDriveBusinessSecret | case | skipped | PASSED |  | OneDrive Business: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_ONEDRIVE_BUSINESS or create 'FileOpsSelfTest OneDrive Business'). |
| FileOperations | Phase16_RemoteOneDriveBusinessSandbox | case | skipped | PASSED |  | OneDrive Business: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_ONEDRIVE_BUSINESS or create 'FileOpsSelfTest OneDrive Business'). |
| FileOperations | Phase16_RemoteSharePointSecret | case | skipped | PASSED |  | SharePoint: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_SHAREPOINT or create 'FileOpsSelfTest SharePoint'). |
| FileOperations | Phase16_RemoteSharePointSandbox | case | skipped | PASSED |  | SharePoint: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_SHAREPOINT or create 'FileOpsSelfTest SharePoint'). |
