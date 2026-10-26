# Final Release connection-diagnostic control

Native **3 passed / 0 failed / 0 skipped**; lookup **237/237**.
Suite 19,485 ms. The exclusively bound, non-listening control takes 11.181745 s.
All five captured failures report Curl 7 / Windows 10061:
one cursor, three bounded size-probe attempts, one namespace attempt.
All three actual HRESULTs are `0x800704C9`; the control verdict is S_OK.

The 233 existing lookup/framing/route checks remain present. The new control
adds four checks; no production timeout, retry, cancellation or mutation authority
changes. Build 04:30, zero warnings/errors, receipt
`10e81ea099b865401a5a72d83377a42534d8087a3dc6011a48c9cbabe17eaf52`.
This focused run overlaps the disjoint-profile Debug build, so its timing is not
a controlled throughput comparison.

See [combined evidence and continuation](../2026-09-05_c0_curl_connect_diagnostics_lookup_debug/README.md).

Run `c0-curl-connect-diagnostics-final-lookup-release-20260905`:

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter C0_CurlEntryLookupTruth -RunId c0-curl-connect-diagnostics-final-lookup-release-20260905
```

Raw evidence is retained in the named governed run under `C:/RedSalamander.Perf/runs/`. Native/root/runner results, traces and entry fingerprints are copied byte-for-byte, with hashes in `file-provenance.json`. `build-provenance.json` binds the test-enabled full-solution receipt. No settings or credentials are archived. Use a new run ID for a repeat.
