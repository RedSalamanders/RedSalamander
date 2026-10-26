# C0 rejected loopback-address experiment

Based on `a46a5a570af09b6cccb8937bcca4609eff2bf3ca`. **C0 remains open.**
The experiment is reverted; no production networking changes are retained.

## Result

Separating native Copy/Move's source and destination listeners onto `127.0.0.2`
and `127.0.0.3` did not cure the post-wide connection failure. Both broad runs
finish all 46 scenarios at **134/142 checks**, with **15 passed / 1 failed / 1
not reached** native cases. The four depth-128 Move routes fail, with four cursor
`CURLE_COULDNT_CONNECT` / Windows 10048 rows per run. Depth-129 refusal cases pass.
The first Move publishes its one file but returns partial-copy; later Move routes
fail before publication. The scenario verdicts remain failed; neither endpoint's
S_OK health nor the increased passing count makes this a successful repair.

| Evidence | Parallel wide | Serial wide | Publications per wide case |
|---|---:|---:|---:|
| [Default-address baseline](../2026-09-05_c0_curl_connect_diagnostics_broad_release/README.md) | 7.988807 s | 25.168114 s | 4,097 |
| This candidate broad run | 9.100050 s | 27.551869 s | 4,097 |
| [Candidate repeat with all-state sample](../2026-09-05_c0_curl_loopback_topology_owned_states_release/README.md) | 9.563265 s | 33.961110 s | 4,097 |

The native Copy workload, 120-second deadline, operation/connect limits, 64-listener
creation cap, ordering, byte/content/terminal assertions and 142-check guard are
unchanged. Both candidate runs retain 13/4 total passive-listener creations and
16,378/16,387 reuses for parallel/serial wide Copy. Other fixtures stay on `.1`.
The experiment also records the two server control ports in `NativeCopy.Begin`;
this instrumentation is reverted with the candidate. It never identifies a
failed secondary socket from a Curl easy handle's last connection metadata.

No build or policy suite overlaps these runtime runs, but each broad run includes
a read-only OS socket query. No clean throughput comparison is claimed.

## Socket observations and limits

An initial owned .NET socket check observed `.2:49437` receiving from `.1:49438`
and `.3:49439` from `.1:49440`. Each listener/client was disposed after the check.
This small probe alone was not used as native evidence.

The first native sample verifies exact executable/PID/start time and observes
client `.1` connected to `.2` and `.3`. It retains 418 rows, but deliberately
filters to loopback peers and therefore omits wildcard BOUND rows. It cannot
rule out retained bound ports. The second sample, triggered by serial-wide
`NativeCopy.Begin`, retains all 1,092 process-owned rows, including 362 BOUND.
It records collection start/end and includes existing closed-control state from
earlier fixtures. Neither sample attributes PID-0 TIME_WAIT rows to the process,
proves an unbounded leak, or establishes the cause of port allocation failure.
Raw rows and exact acquisition metadata are retained in the respective archives.

Windows documents that [implicit-bind connect already supports ephemeral reuse](https://learn.microsoft.com/en-us/windows/win32/winsock/sol-socket-socket-options),
and warns that [TIME_WAIT counts alone do not prove exhaustion](https://learn.microsoft.com/en-us/troubleshoot/windows-client/networking/tcp-ip-port-exhaustion-troubleshooting).
The reviewed Curl 8.21.0 ordinary socket path uses implicit-bind connect when no
local interface/port is configured. There is no basis here for adding reuse/reset
options or changing OS port/TIME_WAIT settings. No such changes were made.

## Invalid predecessor invocation

The previously documented comma-separated four-case filter is rejected by the
FileOps native phase selector before any test case runs. Its [failed invocation](../2026-09-05_c0_curl_loopback_topology_invalid_filter_release/README.md)
is preserved. `SelfTest::CaseFilterMatches` supports comma lists, but
`ResolveRunSelection` does not; do not infer support from the shared matcher.
Use `-CaseFilter C0_` to retain the predecessors in one process until native
selection/listing/host-requirement/reporting behavior has fail-closed tests.

## Provenance and continuation

Test-enabled full-solution Release: 4:27.743, zero warnings/errors; receipt
`8cdbb84fd8bded8dd293577baa25e6a2e6732ca289cfdf1ff04bf323318ae2b9`;
snapshot `d456576b3d205443aec2847953583b8c3fa7f852f8f3237b2073a806515c9932`.
`candidate.patch` preserves the rejected source/spec delta against the base;
`candidate-source-bytes.json` records tested file hashes. The build log retains
its raw and normalized hashes. Source/localization policy passes 197/197 with
zero skips in 19.7607757 s; the original XML is included. Spec inventory has zero
blocking findings. Archive/status changes and the revert follow runtime validation.
There is no Debug candidate run or final-HEAD/Fresh Full claim.

The current C++ matches the base again. The existing Release binaries do not;
rebuild before further qualification. The rejected topology must not silently
become the new baseline. Keep both broad failures and the invalid-filter result.
Next: trace owned socket allocation/close/cache lifetimes at the actual 10048
failure, including BOUND states; independently fix measured fake LIST serialization
cost with byte/fault oracles. Then repeat broad C0, lookup and CRUD in both builds.
The multi-case selector discrepancy needs a focused fix/test if that smaller
predecessor selection is to remain in the workflow. C0's Move-memory/provider
matrix and C1–C6, final Fresh Full and inherited I17 are still open.

Hello remains off only in the five approved isolated profiles. Normal settings,
TLS/SSH policy and configured-server data are unchanged. IMAP verified-TLS and
unique-child live mutations remain pending user choices. Repository-root
`last_run/` and existing sandbox data are untouched; disk-audit findings are not
deleted or hidden. No push is implied.

## Reproduce the archived candidate

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter C0_ -RunId c0-curl-loopback-topology-broad-release-20260905
```

That command requires the exact archived candidate build; use a fresh run ID for
a repeat. Raw artifacts remain under the named governed run. Results/traces and
fingerprints are copied byte-for-byte with `file-provenance.json`. All Curl/C0
metrics are selected in original order; `metric-selection.json` binds raw byte/
row counts, SHA-256, selected line numbers and output hash. No secrets are archived.
