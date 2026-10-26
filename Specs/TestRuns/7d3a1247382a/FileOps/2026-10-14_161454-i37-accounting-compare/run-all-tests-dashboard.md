# Run-AllTests Case Dashboard

- Run: `20261014T161454Z-37216-7ca67aa53530419cbda50bcef20f4872`
- Suite: `Compare`

## Top Slowest Cases

| Harness | Case | Source | Status | Classification | Duration ms | Seed | Attempt |
|---|---|---|---|---|---:|---|---:|
| CompareDirectories | local_search_service_disconnect_falls_back_local_index | case | passed | PASSED | 32561 |  | 1 |
| CompareDirectories | local_search_scan_wide_tree_parallel_walk_name_only | case | passed | PASSED | 8459 |  | 1 |
| CompareDirectories | mtp_mutating_item_commands_time_out | case | passed | PASSED | 8240 |  | 1 |
| CompareDirectories | subdir pending | case | passed | PASSED | 7988 |  | 1 |
| CompareDirectories | search_service_sqlite_cold_start_stale_root_refreshes_before_query | case | passed | PASSED | 5749 |  | 1 |
| CompareDirectories | sqlite_maintenance_busy_checkpoint_and_callback_skip_observability | case | passed | PASSED | 5620 |  | 1 |
| CompareDirectories | mtp_runtime_refresh_defers_when_worker_quarantined | case | passed | PASSED | 5535 |  | 1 |
| CompareDirectories | mtp_menu_and_directory_size_time_out | case | passed | PASSED | 4116 |  | 1 |
| CompareDirectories | mtp_writer_commit_times_out | case | passed | PASSED | 4057 |  | 1 |
| CompareDirectories | search_service_sqlite_delete_burst_maintenance_preserves_query_parity | case | passed | PASSED | 3349 |  | 1 |
| CompareDirectories | test_support_child_process_runner_contract | case | passed | PASSED | 2869 |  | 1 |
| CompareDirectories | decision_cache_eviction_budget_pending_wide_tree | case | passed | PASSED | 2735 |  | 1 |
| CompareDirectories | search_service_sqlite_startup_warmup_failure_status_roundtrip | case | passed | PASSED | 2730 |  | 1 |
| CompareDirectories | decision_cache_eviction_budget_pins_visible | case | passed | PASSED | 2681 |  | 1 |
| CompareDirectories | content_queue_bounded_hi_lo | case | passed | PASSED | 2636 |  | 1 |
| CompareDirectories | invalidate_drops_stale_inflight_scan_writeback | case | passed | PASSED | 2146 |  | 1 |
| CompareDirectories | mtp_hung_device_times_out | case | passed | PASSED | 2067 |  | 1 |
| CompareDirectories | mtp_mutating_create_directory_times_out | case | passed | PASSED | 2056 |  | 1 |
| CompareDirectories | scan_inflight_stamp_guards_restart | case | passed | PASSED | 2038 |  | 1 |
| CompareDirectories | sync_manifest_pending_blocks_and_schedules | case | passed | PASSED | 2012 |  | 1 |

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
| CompareDirectories | local_index_core_sqlite_unrelated_pending_legacy_keeps_requested_ready | case | skipped | PASSED |  | Live journal cursor unavailable; direct SQLite query cutover cannot prove currentness. |
| CompareDirectories | local_index_core_sqlite_prefilter_classifies_name_patterns | case | skipped | PASSED |  | Live journal cursor unavailable; direct SQLite query cutover cannot prove currentness. |
| CompareDirectories | search_service_sqlite_ntfs_traversal_seed_stays_degraded | case | skipped | PASSED |  | NTFS traversal-seed degradation test requires an NTFS-backed self-test root. |
| CompareDirectories | search_service_candidate_impersonation_failure_is_incomplete_warning | case | skipped | PASSED |  | Search service test hooks are unavailable in this build configuration. |
| CompareDirectories | host_fallback_search_remote_ftp_name_only | case | skipped | PASSED |  | FTP: connection profile not found (set REDSALAMANDER_SELFTEST_CONN_FTP or create 'FileOpsSelfTest FTP'). |
| CompareDirectories | local_search_service_root_rejection_does_not_arm_transport_cooldown | case | skipped | PASSED |  | Search service test hooks are unavailable in this build configuration. |
| CompareDirectories | search_service_transient_parent_failure_is_cached_and_warns_batch_and_completion | case | skipped | PASSED |  | Search service test hooks are unavailable in this build configuration. |
