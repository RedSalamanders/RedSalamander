# Run-AllTests Case Dashboard

- Run: `20260712T181655Z-65640-034d58a8c4664073b54c5f4400188407`
- Suite: `Full`

## Top Slowest Cases

| Harness | Case | Source | Status | Classification | Duration ms | Seed | Attempt |
|---|---|---|---|---|---:|---|---:|
| FileOperations | Phase11_BridgeMultiFolderParallelCopyInFlightLines | case | passed | PASSED | 194110 |  | 1 |
| FileOperations | Phase11_BridgePipelineDummyToDummyPerf | case | passed | PASSED | 131610 |  | 1 |
| FileOperations | Phase7_ParallelCopyMoveKnobs | case | passed | PASSED | 95516 |  | 1 |
| FileOperations | Phase5_PreCalcSettingsApplied | case | passed | PASSED | 65578 |  | 1 |
| Commands | cmd_compare_directories_options_long_run_open_close_stays_stable | case | passed | PASSED | 60721 |  | 1 |
| Commands | cmd_compare_directories_options_live_dx_body_interaction | case | passed | PASSED | 55932 |  | 1 |
| FileOperations | Phase8_PerItemOrchestration | case | passed | PASSED | 50563 |  | 1 |
| FileOperations | Phase11_ConnectionOverrideGlobalGate | case | passed | PASSED | 40688 |  | 1 |
| CompareDirectories | local_search_service_disconnect_falls_back_local_index | case | passed | PASSED | 33890 |  | 1 |
| CompareDirectories | search_service_foreground_rejects_second_instance | case | passed | PASSED | 31624 |  | 1 |
| Commands | folderView_perf_overlay_invalidation_stress | case | passed | PASSED | 31554 |  | 1 |
| Commands | cmd_preferences_dialog_category_switches_do_not_churn_tree_host | case | passed | PASSED | 31475 |  | 1 |
| FileOperations | Phase6_ParallelBandwidthThrottleFairness | case | passed | PASSED | 25344 |  | 1 |
| FileOperations | Fairstream_MoveFallbackPreservesUncopiedNewFiles | case | passed | PASSED | 25313 |  | 1 |
| FileOperations | Phase9_PerItemConcurrency | case | passed | PASSED | 25265 |  | 1 |
| FileOperations | Fairstream_StorageKindProbed | case | passed | PASSED | 25250 |  | 1 |
| FileOperations | Phase11_BridgeSingleFolderParallelCopyInFlightLines | case | passed | PASSED | 24625 |  | 1 |
| Commands | folderView_column_widths_audit | case | passed | PASSED | 22486 |  | 1 |
| Commands | folderView_perf_scroll_render_stress | case | passed | PASSED | 19149 |  | 1 |
| FileOperations | Setup | case | passed | PASSED | 17454 |  | 1 |

## Failure And Retry Evidence

| Harness | Case | Source | Status | Classification | Seed | Reason |
|---|---|---|---|---|---|---|
| CompareDirectories | windows_hello_cache | case | skipped | PASSED |  | Windows Hello cache requires initialized host connection UI; host returned ERROR_INVALID_WINDOW_HANDLE (hr=0x80070578). |
| CompareDirectories | oauth_refresh_token_storage | case | skipped | PASSED |  | OAuth refresh token persisted store requires initialized host connection UI; host returned ERROR_INVALID_WINDOW_HANDLE (hr=0x80070578). |
| CompareDirectories | google_drive_cleared_client_id_requires_configuration | case | skipped | PASSED |  | Google Drive clientId gate requires initialized host connection UI; host returned ERROR_INVALID_WINDOW_HANDLE (hr=0x80070578). |
| CompareDirectories | google_drive_connection_requires_refresh_token | case | skipped | PASSED |  | Google Drive refresh-token gate requires initialized host connection UI; host returned ERROR_INVALID_WINDOW_HANDLE (hr=0x80070578). |
| CompareDirectories | onedrive_personal_cleared_client_id_requires_configuration | case | skipped | PASSED |  | OneDrive Personal clientId gate requires initialized host connection UI; host returned ERROR_INVALID_WINDOW_HANDLE (hr=0x80070578). |
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
