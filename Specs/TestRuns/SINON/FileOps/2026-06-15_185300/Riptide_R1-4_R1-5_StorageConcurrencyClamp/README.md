# Riptide R1-4/R1-5 Storage Concurrency Clamp

Scope: prove that the host's storage-aware per-item copy/move budget reaches the local filesystem plugin's recursive directory copy workers.

## RED

- Build: `.\build.ps1 -ProjectName RedSalamander -Configuration Debug`
  - Log: `RED-build-msbuild-20260615_184059_545.log`
  - Result: 0 warnings, 0 errors. This build included the strengthened RED selftest and test-only storage override, but not the production clamp.
- Selftest: `--fileops-selftest --selftest-case=Fairstream_StorageKindProbed --selftest-timeout-multiplier=4`
  - Artifact: `RED_Fairstream_StorageKindProbed/`
  - Result: failed for the intended reason:
    `Storage-kind clamp was not propagated into recursive copy: observed 4 in-flight files with host budget 1.`

## GREEN

- Build: `.\build.ps1 -ProjectName RedSalamander -Configuration Debug`
  - Log: `GREEN-build-msbuild-20260615_184658_711.log`
  - Result: 0 warnings, 0 errors.
- Exact selftest: `Fairstream_StorageKindProbed`
  - Artifact: `GREEN_Fairstream_StorageKindProbed/`
  - Result: passed 3/0/0.
- Regression selftest: `Phase7_CopyItemsSingleFolderRecursiveParallelism`
  - Artifact: `GREEN_Phase7_CopyItemsSingleFolderRecursiveParallelism/`
  - Result: passed 3/0/0; proves `FileSystemOptions.copyMoveMaxConcurrency == 0` still preserves plugin-default recursive parallelism.
- Family selftest: `FileOpsFamily_Fairstream`
  - Artifact: `GREEN_FileOpsFamily_Fairstream/`
  - Result: passed 26/0/0.
- Provider smoke: `FileOps_ProviderCapabilityMatrix`
  - Artifact: `GREEN_FileOps_ProviderCapabilityMatrix/`
  - Result: passed 3/0/0.

## Implementation Notes

- `FileSystemOptions` now carries `copyMoveMaxConcurrency`; it is placed after the existing `uint64_t` so `sizeof(FileSystemOptions)` changes and the existing ABI size checks detect mismatched plugins.
- `Task::InitializeFileSystemOptions(...)` sets bandwidth and, only for PerItem copy/move, the host's `_perItemMaxConcurrencyBudget`.
- The local filesystem plugin clamps copy/move recursive fan-out to `min(plugin config, options.copyMoveMaxConcurrency)` when the host supplies a non-zero limit.
- `Fairstream_StorageKindProbed` still validates storage characteristics, then uses a selftest-only storage override to force a host budget of 1 and asserts live recursive copy never exceeds one in-flight file.
