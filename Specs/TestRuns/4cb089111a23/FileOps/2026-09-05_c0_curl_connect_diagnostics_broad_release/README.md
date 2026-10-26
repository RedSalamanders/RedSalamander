# Broad Release C0 — local address/port failure reproduced

Native result: **15 passed / 1 failed / 1 skipped**; Copy **130/142**.
Lookup is explicitly not reached after Copy failure. All 46 Copy scenarios finish.

The unchanged wide corpus completes with exactly 4,097 files:
parallel 7.988807 s, serial 25.168114 s; 13/4 passive-listener creations respectively.
Four depth-128 Move cases and all four depth-129 refusal cases then fail their
expectations. The first Move returns partial-copy after one publication;
the others return the existing connection-refused HRESULT.

All **eight** new `FileOps.Curl.ConnectionFailure` rows are cursor failures with
Curl code 7, `osErrorAvailable=true`, and **Windows 10048 (WSAEADDRINUSE)**,
not server refusal 10061. This proves local address/port allocation failure on
the affected transport. It does not prove which endpoint lifecycle consumes the
resource. Reported positive ports can identify an existing control connection;
negative ports are unavailable, not failed-socket identity.

This uses the same initial receipt/snapshot as the
[initial control run](../2026-09-05_c0_curl_connect_diagnostics_initial_release/README.md).
The later changes correct the new refusal control and clarify metric interpretation;
there is **no causal repair or subsequent broad green** in this step. The false
success-expectation fields for Move do not by themselves prove a false success report.

The 6,207,313-byte raw stream has 17,506 rows. The projection retains all 1,159
`FileOps.Curl.*` and `fileops.c0.*` rows in original order (465,873 bytes).
`metric-selection.json` binds raw/projected hashes and one-based source positions.
The original aggregate failure is retained, not replaced by focused passes.

See [continuation](../2026-09-05_c0_curl_connect_diagnostics_lookup_debug/README.md).

Run `c0-curl-connect-diagnostics-broad-release-20260905`:

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter C0_ -RunId c0-curl-connect-diagnostics-broad-release-20260905
```

Raw evidence is retained in the named governed run under `C:/RedSalamander.Perf/runs/`. Native/root/runner results, traces and entry fingerprints are copied byte-for-byte, with hashes in `file-provenance.json`. `build-provenance.json` binds the test-enabled full-solution receipt. No settings or credentials are archived. Use a new run ID for a repeat.
