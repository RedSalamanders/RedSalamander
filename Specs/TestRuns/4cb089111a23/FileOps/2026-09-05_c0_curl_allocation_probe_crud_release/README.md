# C0 allocation probe: Release fake FTP CRUD

Native suite: **3 passed / 0 failed / 0 skipped**, 3.3 seconds. This unchanged
read/write/create/delete regression uses the [same Release diagnostic build](../2026-09-05_c0_curl_allocation_probe_lookup_release/README.md).
It runs in a new process after broad C0, so its success does not repair broad order.

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter R0fCurl_FakeFtpReadWriteCreateDelete -RunId c0-curl-allocation-probe-crud-release-20260905
```

Exact results, traces and fingerprints retain raw source paths and hashes.
All Curl/C0 metrics remain ordered and digest-bound. Raw data stays under the
named initialized sandbox run. No build overlaps runtime; no normal settings,
credentials or live server data were changed. C0 and Full qualification remain open.
