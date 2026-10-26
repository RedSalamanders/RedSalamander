# Run-AllTests Case Dashboard

- Run: `20260908T164307Z-104260-ba28d93d5bb041e5900189f2db0b2445`
- Suite: `Full`

## Top Slowest Cases

| Harness | Case | Source | Status | Classification | Duration ms | Seed | Attempt |
|---|---|---|---|---|---:|---|---:|
| FileOperations | Phase11_BridgeMultiFolderParallelCopyInFlightLines | case | passed | PASSED | 195969 |  | 1 |
| FileOperations | C0_CurlCopyTraversalTruth | case | passed | PASSED | 179016 |  | 1 |
| FileOperations | R4T3_WideDirectoryCopyCompletes | case | passed | PASSED | 166156 |  | 1 |
| FileOperations | Phase11_BridgePipelineDummyToDummyPerf | case | passed | PASSED | 128609 |  | 1 |
| Commands | cmd_pane_batchRename_admission_worker_owned_cancellable | case | passed | UNCLASSIFIED_FAILURE | 87150 |  | 1 |
| Commands | cmd_preferences_dialog_themes_tab_traversal_live_dx_interaction | case | failed | UNCLASSIFIED_FAILURE | 70028 |  | 1 |
| FileOperations | Phase7_ParallelCopyMoveKnobs | case | passed | PASSED | 56922 |  | 1 |
| Commands | cmd_compare_directories_options_long_run_open_close_stays_stable | case | passed | UNCLASSIFIED_FAILURE | 54834 |  | 1 |
| Commands | cmd_pane_batchRename_engine_large_preview_perf | case | passed | UNCLASSIFIED_FAILURE | 51144 |  | 1 |
| FileOperations | R4T_DeepTreeCopyCompletesIteratively | case | passed | PASSED | 47859 |  | 1 |
| FileOperations | Phase11_ConnectionOverrideGlobalGate | case | passed | PASSED | 40172 |  | 1 |
| Commands | folderView_perf_focus_selection_state | case | passed | UNCLASSIFIED_FAILURE | 36874 |  | 1 |
| FileOperations | Phase6_PopupSmokeResizeAndPause | case | passed | PASSED | 36016 |  | 1 |
| CompareDirectories | local_search_service_disconnect_falls_back_local_index | case | passed | PASSED | 33800 |  | 1 |
| Commands | cmd_preferences_dialog_category_switches_do_not_churn_tree_host | case | passed | UNCLASSIFIED_FAILURE | 32161 |  | 1 |
| Commands | cmd_pane_rename_folder_prompt_batch_button_opens_batch_rename | case | passed | UNCLASSIFIED_FAILURE | 30159 |  | 1 |
| FileOperations | C0_CurlEntryLookupTruth | case | passed | PASSED | 26562 |  | 1 |
| FileOperations | Phase11_BridgeSingleFolderParallelCopyInFlightLines | case | passed | PASSED | 24953 |  | 1 |
| Commands | folderView_perf_overlay_invalidation_stress | case | passed | UNCLASSIFIED_FAILURE | 23161 |  | 1 |
| Commands | folderView_column_widths_audit | case | passed | UNCLASSIFIED_FAILURE | 23033 |  | 1 |

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
| Commands | cmd_preferences_dialog_themes_page_exposes_live_uia_grid_selection | case | failed | UNCLASSIFIED_FAILURE |  | Preferences Themes DX grid did not expose live UI Automation SelectionPattern after selecting the first row. |
| Commands | cmd_preferences_dialog_themes_tab_traversal_live_dx_interaction | case | failed | UNCLASSIFIED_FAILURE |  | Preferences Themes wrapped theme combo focus target not reached during tab traversal. |
| Commands | cmd_preferences_dialog_hot_paths_live_dx_interaction | case | failed | UNCLASSIFIED_FAILURE |  | Preferences Hot Paths page visible DX edit did not settle to the edited value after live UIA mutation. |
| Commands | cmd_preferences_dialog_keyboard_search_action_updates_dxui_surface | case | failed | UNCLASSIFIED_FAILURE |  | Preferences Keyboard page did not settle before deferred-search validation. |
| Commands | cmd_preferences_dialog_keyboard_export_live_dx_interaction | case | failed | UNCLASSIFIED_FAILURE |  | Failed to invoke the visible Preferences Keyboard Export button through live UIA cancel-path interaction. |
| Commands | cmd_preferences_dialog_keyboard_import_live_dx_interaction | case | failed | UNCLASSIFIED_FAILURE |  | Preferences Keyboard page did not settle to the filtered single-row DX state before import interaction validation. |
| Commands | cmd_pane_find_dialog_result_shortcuts_use_shell_clipboard_and_file_actions | case | failed | UNCLASSIFIED_FAILURE |  | Find result context menu selection Clipboard Copy did not expose readable menu state. |
| Commands | cmd_pane_find_dialog_command_enablement_matches_idle_running_and_selection_states | case | failed | UNCLASSIFIED_FAILURE |  | Failed to select 'Z:\RedSalamander.Perf\runs\20260908T164307Z-104260-ba28d93d5bb041e5900189f2db0b2445\artifacts\selftest\last_run\commands\work\find_dialog_command_enablement\sub' for command-enablement test. |
| Commands | cmd_pane_find_dialog_action_buttons_activate_expected_commands | case | failed | UNCLASSIFIED_FAILURE |  | Find button delivered click did not add the expected initial result. |
| Commands | cmd_compare_directories_options_theme_cycle_keeps_surface_legible | case | failed | UNCLASSIFIED_FAILURE |  | Compare Directories options ignore-files edit did not retake focus after the dark theme update. |
| Commands | cmd_pane_selection_select_next_keeps_navigation_shell_stable | case | failed | UNCLASSIFIED_FAILURE |  | Failed to capture the baseline navigation-view state before select-next shell-stability validation. |
| Commands | cmd_pane_filter_dialog_keeps_navigation_shell_stable | case | failed | UNCLASSIFIED_FAILURE |  | Navigation shell did not settle during Pane Filter enable; focusTarget=0, editMode=no, historyVisible=no, suggestVisible=no, popupVisible=no, childWindows=0, currentPath='Z:\RedSalamander.Perf\runs\20260908T164307Z-104260-ba28d93d5bb041e5900189f2db0b2445\artifacts\selftest\last_run\commands\work\pane_filter_nav_shell_aa6cfa8d-b9ee-4ab7-afd3-dfe21ec6b2e9', historyCount=20/20, refreshCount=45/45, itemCount=2/2, selectedCount=1/1, focusedItem='a.txt', focusedFolderView=0x0, expectedFolderView=0x1D20470, filterEnabled=true/true, filterText='*.txt'/'*.txt'. |
| Commands | cmd_pane_navigation_rename_prompt_keeps_navigation_shell_stable | case | failed | UNCLASSIFIED_FAILURE |  | Navigation shell did not restore cleanly after rename confirm; focusTarget=0, editMode=no, historyVisible=no, suggestVisible=no, popupVisible=no, childWindows=0, currentPath='Z:\RedSalamander.Perf\runs\20260908T164307Z-104260-ba28d93d5bb041e5900189f2db0b2445\artifacts\selftest\last_run\commands\work\rename_nav_shell_f5c5408d-28b1-472b-a62d-1ad9051df51b', panePath='Z:\RedSalamander.Perf\runs\20260908T164307Z-104260-ba28d93d5bb041e5900189f2db0b2445\artifacts\selftest\last_run\commands\work\rename_nav_shell_f5c5408d-28b1-472b-a62d-1ad9051df51b', historyCount=20, refreshCount=50, itemCount=1, selectedCount=0, expectedSelectedCount=0, focusedItem='renamed_from_shell_stability.txt', focusedFolderView=0x0, expectedFolderView=0x1D20470, win32Focus=0x0, active=0x0, foreground=0x4053C, foregroundPid=21032, mainWindow=0x1172AFC. |
| Commands | cmd_app_managePlugins_keeps_navigation_shell_stable | case | failed | UNCLASSIFIED_FAILURE |  | App Manage Plugins should restore folder-view focus to the active pane after close. |
| Commands | cmd_viewer_space_layout_20k_visible_optin | case | skipped | UNCLASSIFIED_FAILURE |  | Set REDSALAMANDER_VIEWERSPACE_20K_LAYOUT=1 to run the opt-in ViewerSpace 20k layout scenario. |
| Commands | viewer_space_perf_large_optin | case | skipped | UNCLASSIFIED_FAILURE |  | Set REDSALAMANDER_VIEWERSPACE_LARGE_PERF=1 to run the opt-in ViewerSpace large perf scenario. |
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
