# EG-S3-02 exact S3 source-revision validation

- Date: 2026-08-04
- Machine profile: `7d3a1247382a`
- Source fingerprint at capture: `6aecfde8e7c09e64d48f0c2a86b6d189b67e0e7f` plus the uncommitted EG-S3-02 implementation under test
- Build: test-enabled x64 Release
- Provider build: `.build/logs/msbuild-20260804_102212_897.log` - 0 warnings, 0 errors
- Contract-test build: `.build/logs/msbuild-20260804_101929_591.log` - 0 warnings, 0 errors
- Focused entry point: `RedSalamanderS3DirectoryTransferProbeSelfTests`
- Result: five runs, each 131 passed / 0 failed
- Raw bounded metrics: `perf_metrics.jsonl`

The pre-change failure-first fake graph reported 46 passed / 76 failed. It replaced revision A before small copy,
between multipart parts, before relay read, and before MOVE deletion. Same-size and changed-size replacements were
covered for COPY and MOVE in versioned and unversioned modes. The repaired matrix copies only captured A, retains B,
deletes only captured version A for a versioned MOVE, and fails closed on an unversioned mismatch. Cancellation and a
later conditional-delete failure also prove destination rollback and source preservation.

## Protected performance scenario

The same test-enabled Release binary copies 16 objects under fixed 3-ms fake metadata latency. The unsafe baseline
uses a test-only switch that removes source revision pinning; the production candidate performs one identity HEAD and
one conditional source read per object. Five same-process pairs produced:

| Metric | Unsafe baseline | Pinned candidate |
|---|---:|---:|
| Identity probes | 0 | 16 |
| Conditional source reads | 0 | 16 |
| Median duration | 536,572 us | 786,180 us |
| Median added duration | - | 249,362 us (+46.52%) |

The candidate remains below the deterministic 175% end-to-end gate. This is request-shape and fixed-latency fake
transport evidence; it is not a live S3 throughput claim.
