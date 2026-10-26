# Cinderstar transactional Curl writer validation

- Date: 2026-08-09
- Machine profile: `4cb089111a23`
- Source fingerprint at capture: `e0de6582ee82eaa5bb0711260d04d5b6642a9660` plus the uncommitted Cinderstar package 2 implementation under test
- Build: test-enabled x64 Release
- Provider build: `.build/logs/msbuild-20260809_111152_188.log` - 0 warnings, 0 errors
- Contract-test build: `.build/logs/msbuild-20260809_111223_847.log` - 0 warnings, 0 errors
- Focused entry point: `RedSalamanderCurlDebugSelfTests`
- Result: five runs, each 130 passed / 0 failed
- Raw bounded metrics: `perf_metrics.jsonl`

The real loopback FTP fixture exercises the public writer path. Bytes remain in a local delete-on-close file until
`Commit`, then upload to a unique remote sibling, pass exact `SIZE` verification, preserve an overwrite target under
a rollback sibling, and promote. The matrix covers successful overwrite, abandoned writer, short-success upload,
promotion failure, no-overwrite fail-closed, IMAP fail-closed, strong owner lifetime, and unload quiet point.

## Bounded evidence

| Scenario | Median | Exact endpoint commands | Result |
|---|---:|---:|---|
| Transactional 4,099-byte overwrite | 5,000 us | 27 | `S_OK`, exact final bytes, no siblings |
| Short-success upload (2,049 bytes retained) | 3,000 us | 15 | `ERROR_PARTIAL_COPY`, sentinel retained, stage removed |
| Promotion failure | 5,000 us | 35 | failure, sentinel restored, stage/backup removed |
| Abandoned before Commit | 0 us | 0 | local stage discarded, endpoint untouched |

The unsafe before state streamed directly to the final name and therefore had no staging, size-proof, rollback, or
cleanup transaction to time comparably. Its request-shape baseline was a single final-path upload beginning before
`Commit`; the candidate deliberately adds bounded safety requests. This archive makes no throughput-improvement claim.
It proves bounded latency/commands and mutation safety on a fake loopback transport, not live FTP/SFTP performance.
