# Curl source-commitment and reader-integrity validation

- Date: 2026-08-16
- Machine profile: `4cb089111a23`
- Source base at capture: `e9c76e59aa2b2993170ea1e6c88e7d15362954c7` plus the uncommitted Package B remediation under test
- Build: test-enabled x64 Release
- FileSystemCurl build: `.build/logs/msbuild-20260816_085633_814-pid9952-ad3fc4dc.log` — 0 warnings, 0 errors; receipt `47bd49a72bf0bb552e1be107ee8fa370944d3e426950e96b09e666fbcf10838a`
- PluginContractTests build: `.build/logs/msbuild-20260816_085716_940-pid9952-2a9f6359.log` — 0 warnings, 0 errors; receipt `375cf500f5765eaaf32e102158442097a5d4cb9e7998114d460a9e51c316beaf`
- Focused entry point: `.build\x64\Release\PluginContractTests.exe --curl-selftests`
- Result: five separate processes, each 199 passed / 0 failed with a clean FileSystemCurl unload quiet point
- Raw bounded metrics: `perf_metrics.jsonl` (140 valid JSONL rows, 74,467 bytes, SHA-256 `d40e99beeb400d869998c4392094c701a8c965c903297d4c430fd9043c7a9d13`)

The deterministic loopback matrix protects provider-native Curl Copy/Move and direct-reader source commitments. It covers targeted-size precedence over stale listings, explicit unknown-size policy, authentication and disappearance propagation, exact/short/overlong/zero bodies, single/batch/recursive operations, destructive recursive-Move preflight, FTP `REST` seeking, failed-generation restart, pre-mutation command counts, and unload safety.

## Bounded evidence

| Metric | Five Release samples | Median | Stable work/result |
|---|---|---:|---|
| `FileOps.Curl.SourceCommitmentGuard` duration | 4,000; 6,000; 6,000; 5,000; 6,000 us | 6,000 us | 22 source commands, 4,098 deliberately short bytes delivered, `0x8007012B` (`ERROR_PARTIAL_COPY`) each run |

The metric measures the paired known-size short-body Copy and Move guard. Every run rejects the clean protocol completion before destination `STOR` or source `DELE`; the full 199-case matrix proves the broader behavior. Timing is loopback and millisecond-quantized (`elapsedMs * 1000`), so it is a bounded safety-cost baseline rather than a live FTP/SFTP throughput measurement.

There is no apples-to-apples historical `FileOps.Curl.SourceCommitmentGuard` baseline. The older transactional-writer archive measures publication and rollback, not source-commitment preflight or read termination, and is therefore not used as a before/after comparison. This archive makes no speedup claim.
