# R0b S3 exact-generation Delete baseline

- Machine profile: `4cb089111a23`
- Source commit: `cd9d5f05` (RED harness, before the R0b behavior change)
- Build: test-enabled x64 Release; FileSystemS3 receipt `6b667430d3ffd390f90d8cf86b45533b0cfb366ca388ab7c11e5a626fb43da3b`; PluginContractTests receipt `27b6ee9bf443315f8b08966f58226f454199ddbc55eb0d9ccda9b3f135716e72`
- Command: `PluginContractTests.exe --s3-r0b-delete-selftests`
- Scenario: five independent processes, each deleting 4,096 listed 4-KiB fake S3 objects
- Metric: `FileOps.S3.VirtualFolderDelete.SelfTestUs`
- Samples: `5,623`, `5,790`, `5,037`, `6,334`, `6,051` us
- Nearest-rank p95: `6,334` us
- Request shape: 4,096 delete attempts, `0` conditional requests, two listing passes, zero reported residuals

The focused assertion is intentionally RED (`passed=2, failed=1`): the old ordinary recursive route discarded every listed ETag before mutation. The operation completes and emits its timing before that assertion.
