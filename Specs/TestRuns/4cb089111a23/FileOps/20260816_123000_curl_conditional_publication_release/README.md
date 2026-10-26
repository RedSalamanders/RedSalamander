# Curl conditional-publication and identity-safe-recovery validation

- Date: 2026-08-16
- Machine profile: `4cb089111a23`
- Source base at capture: `e9c76e59aa2b2993170ea1e6c88e7d15362954c7` plus the uncommitted Package C remediation under test
- Build: test-enabled x64 Release
- FileSystemCurl build: `.build/logs/msbuild-20260816_122832_700-pid71284-188583c1.log` — 0 warnings, 0 errors; receipt `c4383f8a109f7f74fb007bff9862cb7ed6a7ed8ee66b5c4478a359406116456e`
- PluginContractTests build: `.build/logs/msbuild-20260816_122921_122-pid71284-bbd6fde7.log` — 0 warnings, 0 errors; receipt `136b4b668cbc0a24bf5f2d222e0c9ae5fa138c9c3d6796d3af17af73c71592db`
- Focused entry point: `.build\x64\Release\PluginContractTests.exe --curl-selftests`
- Result: five separate processes, each 211 passed / 0 failed with a clean FileSystemCurl unload quiet point
- Raw bounded metrics: `perf_metrics.jsonl` (175 valid JSONL rows, 96,609 bytes, SHA-256 `37dfd98b361f4febda93bf856b76b5bd1c4a99e179dafeb8447e219ff1db1550`)

The deterministic loopback matrix protects FTP/SFTP/SCP capability truth, direct no-overwrite admission, and recovery after uncertain remote publication. It covers single and batch Copy, Move, and Rename rejection before endpoint work; late destination creation; post-publication replacement; source-delete failure; writer-promotion failure; concurrent recursive children; artifact preservation; and unload safety. The broader 211-case matrix also retains the accepted Package B source-commitment and reader-integrity coverage.

## Bounded evidence

| Metric | Five Release samples | Median | Stable work/result |
|---|---|---:|---|
| `FileOps.Curl.ConditionalPublicationGuard` duration | 0; 0; 0; 0; 0 us | 0 us | Six direct no-overwrite shapes, zero endpoint commands, zero race injections, `0x80070032` (`ERROR_NOT_SUPPORTED`) each run |
| `FileOps.Curl.IdentitySafeRecoveryGuard` duration | 20,000; 22,000; 18,000; 20,000; 21,000 us | 20,000 us | 103 deterministic loopback commands, three preserved final/backup/tree artifacts, `0x8007012B` (`ERROR_PARTIAL_COPY`) each run |

The conditional-publication metric is zero because admission rejects before endpoint work and the harness records millisecond-quantized elapsed time. The recovery metric combines three deterministic loopback failure scenarios, so it is a bounded safety-cost measurement rather than FTP/SFTP throughput. There is no apples-to-apples historical metric for these guards; the earlier Curl source-commitment and transactional-writer archives protect different behavior and are not used as baselines. This archive is same-machine evidence but makes no speedup or live-endpoint claim.
