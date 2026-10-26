# C0 listing candidate: Debug basic fake FTP CRUD

Native suite: **3 passed / 0 failed / 0 skipped**, 2.8 seconds. This is the
unchanged basic read/write/create/delete regression against the same
[test-enabled Debug candidate and receipt](../2026-09-05_c0_curl_listing_lookup_debug/README.md).
It checks owned loopback fixture behavior, not live FTP/SFTP/SCP servers.

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter R0fCurl_FakeFtpReadWriteCreateDelete -RunId c0-curl-listing-candidate-crud-debug-20260905
```

Exact results, traces and fingerprints retain their raw source paths and hashes
in `file-provenance.json`; the complete Curl/C0 metric projection is ordered and
digest-bound. Raw data remains under that initialized sandbox run. Reproduction
requires the matching candidate and a fresh run ID. No build overlaps runtime;
normal settings, credentials and live endpoint data are unchanged. C0 remains open.
