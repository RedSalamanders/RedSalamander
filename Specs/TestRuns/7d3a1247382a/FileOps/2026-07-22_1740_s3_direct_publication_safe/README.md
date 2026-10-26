# EG-S3-01 conditional direct-publication validation

- Date: 2026-07-22
- Machine profile: `7d3a1247382a`
- Source fingerprint at capture: `3ab6069f1d23` plus the uncommitted EG-S3-01 implementation under test
- Build: test-enabled x64 Release
- Full-solution rebuild: `.build/logs/msbuild-20260722_172620_876.log` — 10 existing solution-wide warnings, 0 errors
- Command: `.build\x64\Release\PluginContractTests.exe`
- Result: exit 0; complete executable passed; S3 directory-transfer probes reported 37 passed / 0 failed
- Source contracts: 149 passed / 0 failed
- Production tests-disabled Release provider build: `.build/logs/msbuild-20260722_174751_212.log` — 0 warnings, 0 errors
- Final Debug/x64 rebuild: `.build/logs/msbuild-20260722_185201_566.log` — 0 warnings, 0 errors
- Debug provider contracts: complete `PluginContractTests.exe` passed
- Debug FileOps: 111 passed / 0 failed / 20 expected environment skips; zero failure classifications
- Raw metrics: `perf.jsonl`

The repository-wide Full attempt built Debug with 0 warnings/errors and completed Compare at 227/0 with 31 expected
environment skips. Commands had recorded 371 passed / 0 failed when the outer one-hour command cap interrupted the
next UI case, before FileOps and standalone aggregation. That attempt is inconclusive rather than a product failure;
the subsequent clean rebuild, complete Debug provider contracts, focused FileOps aggregate, Release provider
contracts, and source contracts are the authoritative closeout gates above.

## Correctness proof

The fake S3 graph creates the competing destination inside publication, after whole-plan and just-in-time probes
reported it missing. Small `CopyObject`, forced multipart completion, and relay upload are covered for COPY and MOVE.
Every conditional conflict preserves the competing destination bytes, and every failed MOVE preserves its source.
The matrix also proves multipart abort, an endpoint that explicitly rejects conditional relay publication, and late
conflict re-entry through Skip and Overwrite.

## Protected performance scenario

The same binary copies 16 objects without conflicts under fixed 3-ms publication latency. A test-only scope supplies
the unsafe unconditional baseline; production behavior is the conditional candidate.

| Run | Publication time | Requests | Conditional requests | Relay fallbacks | Conflicts |
|---|---:|---:|---:|---:|---:|
| Unsafe baseline | 251,041 us | 16 | 0 | 0 | 0 |
| Conditional candidate | 253,786 us | 16 | 16 | 0 | 0 |

The header changes the request condition without adding a request. Measured overhead is 2,745 us (1.09%), within the
30% deterministic timing-noise gate. This is request-shape and fixed-latency selftest evidence, not live S3 endpoint
throughput evidence.
