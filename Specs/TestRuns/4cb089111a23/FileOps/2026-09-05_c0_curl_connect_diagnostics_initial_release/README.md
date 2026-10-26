# Initial connection-diagnostic control — failure retained

Native result: **2 passed / 1 failed / 0 skipped**. Lookup: **234/237**. The existing
233 checks and new endpoint setup pass; all three new refusal-result checks fail.
The initial control used 500 ms connect / 2 s operation limits and did not record
each actual HRESULT. Its aggregate control metric fails after 3.583480 s.

An independent exclusively bound, non-listening loopback socket probe reports
Windows 10061 after 2,064 ms. That supports deadline truncation as the initial
control problem; it is not evidence that the product misclassified an observed
refusal. The final control uses 5 s / 7 s limits and records all actual HRESULTs.
Both configurations then pass 237/237. Only this new control's deadlines changed.

Build receipt: `51b4b28528b73db86aabd06066b36c4331f15755613245f41f63ac2173dbf4b2`.
Snapshot: `54236672731926baacdec359ba23952099e93c4b536d822fc11f7e0d09dd15de`.
Build 04:04, **one MSB3061 warning**, zero errors: Google Drive Clean cannot delete
shared `Plugins/libcrypto-3-x64.dll`. The warning-bearing log is retained.
Initial source/resource XML: 197 passed, zero failed/skipped, 19.1766748 s.
The first policy launch was denied creating its scratch directory; the authorized
retry used the same owned run ID before any completed result existed.

See [final evidence and continuation](../2026-09-05_c0_curl_connect_diagnostics_lookup_debug/README.md).

Run `c0-curl-connect-diagnostics-lookup-release-20260905`:

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter C0_CurlEntryLookupTruth -RunId c0-curl-connect-diagnostics-lookup-release-20260905
```

Raw evidence is retained in the named governed run under `C:/RedSalamander.Perf/runs/`. Native/root/runner results, traces and entry fingerprints are copied byte-for-byte, with hashes in `file-provenance.json`. `build-provenance.json` binds the test-enabled full-solution receipt. No settings or credentials are archived. Use a new run ID for a repeat.
