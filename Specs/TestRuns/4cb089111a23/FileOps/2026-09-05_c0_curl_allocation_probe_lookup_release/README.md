# C0 allocation probe: Release control and build provenance

Based on `dbc75c7d9ab716991f79832ac6f9e7e94c135ce8`. Native suite: **3 passed,
0 failed, 0 skipped**, 20.2 seconds; exact lookup **250/250**. The new control
obtains a nonzero ephemeral port (56115) and preserves the deliberately seeded
Winsock error 10050. The prior 249 checks remain unchanged.

The helper creates one owned TCP socket, binds wildcard address/port zero, observes
the assigned port, closes, and restores the caller's Winsock error on every return.
It never listens, connects, alters socket options or changes machine settings.
Outside the control it runs only after a captured 10048 error in test-enabled
builds. It cannot grant retry/mutation authority or change the operation HRESULT.
The observed port is already released, not a reservation for a later connection.

Release test-enabled full-solution build: **248,377 ms**, zero warnings/errors.
Receipt `ae9ab21d32c95b6f15717d5fbea59ea82a90555e00a195eeed20197b65615da7`;
snapshot `3a6f609f3361388a361e02c57d3f1dea590403b3ddbb414f3aa34908be75a74d`.
Source/localization policy: **197/197**, zero skips, 19.7029931 seconds. Policy
overlaps the build, not runtime. `tested.patch` and `tested-source-bytes.json`
preserve the tested delta before final evidence-only plan/attribute updates.

Read-only `Get-NetTCPSetting` output is retained exactly. All five concrete
templates report dynamic start/count 49152/16384 and auto-reuse start/count 0/0;
Automatic reports unavailable fields. Microsoft documents that a zero auto-reuse
range disables that feature even with the socket option enabled. This contextual
fact does not prove exhaustion. Neither the range nor any other system setting
was changed. Sources: [TCP setting semantics](https://learn.microsoft.com/en-us/powershell/module/nettcpip/get-nettcpsetting?view=windowsserver2025-ps),
[socket option semantics](https://learn.microsoft.com/en-us/windows/win32/winsock/sol-socket-socket-options).

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter C0_CurlEntryLookupTruth -RunId c0-curl-allocation-probe-lookup-release-20260905
```

Raw data remains under that initialized sandbox run (policy/settings in the named
policy run). Results, traces, fingerprints, policy XML and settings are exact
copies with source hashes. All Curl/C0 metric rows remain ordered, with raw
bytes/rows/hash and selected-line/hash binding. Build-log normalization is explicit.
Normal connections, credentials and live endpoint data are untouched. This is
diagnostic coverage, not C0 or Full closeout.
