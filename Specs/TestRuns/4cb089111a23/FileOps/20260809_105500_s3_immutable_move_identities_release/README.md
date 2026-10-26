# Cinderstar immutable S3 MOVE identity validation

- Date: 2026-08-09
- Machine profile: `4cb089111a23`
- Source fingerprint at capture: `e0de6582ee82eaa5bb0711260d04d5b6642a9660` plus the uncommitted Cinderstar package 1 implementation under test
- Build: test-enabled x64 Release
- Provider build: `.build/logs/msbuild-20260809_105127_128.log` - 0 warnings, 0 errors
- Contract-test build: `.build/logs/msbuild-20260809_105235_309.log` - 0 warnings, 0 errors
- Focused entry point: `RedSalamanderS3DirectoryTransferProbeSelfTests`
- Result: five runs, each 139 passed / 0 failed
- Raw bounded metrics: `perf_metrics.jsonl`

The deterministic fake graph covers immutable final-publication and backup identities, exact source delete-marker
identity, versioned source-history hiding, source replacement, destination replacement after publication, cancellation,
later source-delete failure, and exact rollback cleanup. Assertions compare visible keys and exact version sets. Across
five repetitions, the protected 16-object scenarios captured all 16 publication identities, the versioned MOVE
scenario captured exactly one delete-marker identity, and rollback/cleanup scenarios removed only operation-owned
revisions.

## Protected performance scenario

The 2026-08-04 source-pinning archive is the before point for the same 16-object fixed-3-ms fake-metadata scenario.
It proved source identity pinning only; it did not prove immutable destination or delete-marker rollback identities.
The current run adds those identity assertions and counters without adding a provider request to successful
publication because the identity comes from the existing SDK response.

| Metric | 2026-08-04 before | Cinderstar candidate |
|---|---:|---:|
| Provider-contract checks per repeat | 131 | 139 |
| Immutable publication identities per protected run | not recorded | 16 / 16 |
| Versioned source delete-marker identities | not recorded | 1 / 1 |
| Median pinned-candidate duration | 786,180 us | 775,035 us |
| Change | - | -11,145 us (-1.42%) |

The current same-binary unsafe/pinned pair remains inside its 175% gate: `527,169 us -> 775,035 us` (+46.98%).
This is request-shape and fixed-latency fake-transport evidence, not a live S3 throughput claim.
