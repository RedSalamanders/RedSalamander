# R0b S3 exact-generation Delete candidate

- Machine profile: `4cb089111a23`
- Source commit: `e172269d` (R0b exact-generation implementation)
- Build: test-enabled x64 Release; FileSystemS3 receipt `af853df54434cbfc16c1bd95fb6559e4e31066612ecc945e4fb722707f27191f`; PluginContractTests receipt `622ad0c7e2347ed34c692834744e09c5b1ae86e2bd829ad5a6ba20c7420704f4`
- Command: `PluginContractTests.exe --s3-r0b-delete-selftests`
- Scenario: five independent processes, each deleting 4,096 listed 4-KiB fake S3 objects
- Metric: `FileOps.S3.VirtualFolderDelete.SelfTestUs`
- Samples: `5,290`, `6,094`, `5,882`, `6,547`, `6,613` us
- Nearest-rank p95: `6,613` us
- Request shape: 4,096 delete attempts, 4,096 ETag-conditioned requests, two listing passes, zero residuals, one live request at a time

The focused contract is GREEN (`passed=3, failed=0`). Candidate p95 is 279 us / 4.40% above baseline and passes the safety-stabilization ceiling `max(1.50 * 6,334, 6,334 + 50,000) = 56,334 us`; every sample is below 2,000,000 us. This fake-graph evidence protects CPU/memory/request shape and does not claim live-AWS throughput.
