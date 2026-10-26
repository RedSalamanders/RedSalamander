# C0 allocation probe: Debug fake FTP CRUD

Native suite: **3 passed / 0 failed / 0 skipped**, 5.6 seconds. The unchanged
read/write/create/delete regression uses the [same Debug diagnostic build](../2026-09-05_c0_curl_allocation_probe_lookup_debug/README.md)
after the exact lookup run. Neither this isolated success nor the successful
allocation control repairs the preserved broad-process failure.

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter R0fCurl_FakeFtpReadWriteCreateDelete -RunId c0-curl-allocation-probe-crud-debug-20260905
```

Exact result/trace/fingerprint copies retain source paths and hashes. All Curl/C0
metrics remain ordered and digest-bound. Raw data remains under the governed run;
its 133 disk-audit findings are not removed. No build overlaps runtime, and normal
settings, credentials and live server data are untouched. C0 remains open.
