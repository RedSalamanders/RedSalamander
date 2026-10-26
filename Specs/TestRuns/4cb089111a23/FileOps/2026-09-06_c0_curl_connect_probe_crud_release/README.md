# C0 connection probes: Release fake FTP CRUD

Native suite **3 passed / 0 failed / 0 skipped**, 2.9 seconds. The unchanged
read/write/create/delete regression uses the [same Release diagnostic build](../2026-09-06_c0_curl_connect_probe_lookup_release/README.md),
in a fresh process after broad C0. This does not repair the broad-order failure.

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter R0fCurl_FakeFtpReadWriteCreateDelete -RunId c0-curl-connect-probe-crud-release-20260906
```

Exact results, traces and fingerprints retain source hashes. Ordered Curl/C0
metrics remain digest-bound. Raw data and all 137 disk-audit findings remain under
the governed run. No build overlaps runtime; no normal settings, credentials or
real endpoint data changed. C0 remains open.
