# Run-AllTests Case Dashboard

- Run: `20260923T113614Z-77640-e22df5f2325745d1839d9a66e31aa8df`
- Suite: `FileOps`

## Top Slowest Cases

| Harness | Case | Source | Status | Classification | Duration ms | Seed | Attempt |
|---|---|---|---|---|---:|---|---:|
| FileOperations | Phase11_BridgeMultiFolderParallelCopyInFlightLines | case | passed | PASSED | 194078 |  | 1 |
| FileOperations | Phase11_BridgePipelineDummyToDummyPerf | case | passed | PASSED | 128344 |  | 1 |
| FileOperations | R4T_DeepTreeCopyCompletesIteratively | case | passed | PASSED | 47188 |  | 1 |
| FileOperations | C0_CurlCopyTraversalTruth | case | passed | PASSED | 45016 |  | 1 |
| FileOperations | Phase11_ConnectionOverrideGlobalGate | case | passed | PASSED | 39875 |  | 1 |
| FileOperations | Phase7_ParallelCopyMoveKnobs | case | passed | PASSED | 37797 |  | 1 |
| FileOperations | Phase11_BridgeSingleFolderParallelCopyInFlightLines | case | passed | PASSED | 24938 |  | 1 |
| FileOperations | Phase6_ParallelBandwidthThrottleFairness | case | passed | PASSED | 18969 |  | 1 |
| FileOperations | Phase8_PerItemOrchestration | case | passed | PASSED | 18891 |  | 1 |
| FileOperations | C0_CurlMovePreflightTruth | case | passed | PASSED | 17781 |  | 1 |
| FileOperations | Phase6_PopupSmokeResizeAndPause | case | passed | PASSED | 17250 |  | 1 |
| FileOperations | Phase9_PerItemConcurrency | case | passed | PASSED | 17078 |  | 1 |
| FileOperations | Phase11_ConnectionOverrideClamp | case | passed | PASSED | 15000 |  | 1 |
| FileOperations | R3_3_PublicationFaultMatrix | case | passed | PASSED | 14703 |  | 1 |
| FileOperations | R4T3_WideDirectoryCopyCompletes | case | passed | PASSED | 13015 |  | 1 |
| FileOperations | R4A02_RunConcurrentSameDestinationStaysSafe | case | passed | PASSED | 12954 |  | 1 |
| FileOperations | Fairstream_StorageKindProbed | case | passed | PASSED | 12547 |  | 1 |
| FileOperations | Phase11_CrossFileSystemBridge | case | passed | PASSED | 11125 |  | 1 |
| FileOperations | Fairstream_MovedTreeKeepsLiteralLinks | case | passed | PASSED | 11031 |  | 1 |
| FileOperations | Phase7_RecycleBinBatchDelete | case | passed | PASSED | 10391 |  | 1 |

## Failure And Retry Evidence

| Harness | Case | Source | Status | Classification | Seed | Reason |
|---|---|---|---|---|---|---|
| FileOperations | Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker | case | skipped | PASSED |  | FileSystem shared-scheduler shutdown selftests use a Debug-only plugin export; Debug coverage owns this proof. |
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
