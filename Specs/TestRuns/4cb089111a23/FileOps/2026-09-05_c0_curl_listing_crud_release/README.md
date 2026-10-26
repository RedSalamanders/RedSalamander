# C0 listing candidate: Release basic fake FTP CRUD

Native suite: **3 passed / 0 failed / 0 skipped**, 3.5 seconds. This is the
unchanged basic read/write/create/delete regression against the
[test-enabled Release candidate](../2026-09-05_c0_curl_listing_lookup_release/README.md),
run after the broad failure in a new process. It is not a broad-order repair or
live-server qualification.

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter R0fCurl_FakeFtpReadWriteCreateDelete -RunId c0-curl-listing-candidate-crud-release-20260905
```

Exact results, traces and fingerprints retain raw source paths and hashes in
`file-provenance.json`; all Curl/C0 metrics remain ordered and digest-bound.
Raw data stays under that initialized sandbox run. Reproduction requires the
matching build and a fresh run ID. There is no overlapping build, policy check or
socket sampling, and no change to normal settings, credentials or live endpoints.
