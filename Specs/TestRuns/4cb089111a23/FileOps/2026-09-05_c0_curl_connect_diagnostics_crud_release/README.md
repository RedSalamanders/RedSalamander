# Release successful-path FTP CRUD regression

Native **3 passed / 0 failed / 0 skipped** (Setup, CRUD, Cleanup).
The ordinary loopback read/write/create/delete scenario passes using the same
receipt as the [final Release lookup](../2026-09-05_c0_curl_connect_diagnostics_lookup_release/README.md).
These are test-owned local endpoints, not newly authorized configured-server mutations.
This focused pass does not replace the broad C0 failure.

See [combined evidence and continuation](../2026-09-05_c0_curl_connect_diagnostics_lookup_debug/README.md).

Run `c0-curl-connect-diagnostics-crud-release-20260905`:

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter R0fCurl_FakeFtpReadWriteCreateDelete -RunId c0-curl-connect-diagnostics-crud-release-20260905
```

Raw evidence is retained in the named governed run under `C:/RedSalamander.Perf/runs/`. Native/root/runner results, traces and entry fingerprints are copied byte-for-byte, with hashes in `file-provenance.json`. `build-provenance.json` binds the test-enabled full-solution receipt. No settings or credentials are archived. Use a new run ID for a repeat.
