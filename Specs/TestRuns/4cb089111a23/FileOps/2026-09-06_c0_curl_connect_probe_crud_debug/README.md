# C0 connection probes: Debug fake FTP CRUD

Native suite **3 passed / 0 failed / 0 skipped**, 3.4 seconds. The unchanged
read/write/create/delete regression uses the [same Debug diagnostic build](../2026-09-06_c0_curl_connect_probe_lookup_debug/README.md),
after the exact lookup control. Its fresh-process pass does not repair broad C0.

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter R0fCurl_FakeFtpReadWriteCreateDelete -RunId c0-curl-connect-probe-crud-debug-20260906
```

Exact results, traces and fingerprints retain source hashes. Ordered Curl/C0
metrics remain digest-bound. Raw data and 139 disk-audit findings remain in the
governed run. No build overlaps runtime. Normal settings, credentials and real
endpoint data are unchanged. C0 remains open.
