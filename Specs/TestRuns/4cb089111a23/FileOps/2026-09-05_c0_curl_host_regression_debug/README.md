# Curl host workflow after native Delete repairs

- Source: `58f22d0e` plus the current C0 receipt and native Curl Delete changes.
- Debug build receipt:
  `2940050c3e6a62468306357bb7086ba48083ffdf1894da21da7a0f17beb0731a`;
  zero warnings/errors.
- Command: `Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter R0fCurl_FakeFtpReadWriteCreateDelete -FailFast -RunId c0-curl-host-regression-debug-20260905`.
- Result: 3/3 passed, zero skips (setup, actual host workflow, cleanup).
- Raw results, trace, aggregate, and full 715,668-byte metric stream are unchanged
  copies from the run-owned `artifacts/selftest/last_run` directory.

The existing host test exercises copy to/from fake FTP, destination replacement,
directory creation/rename, and deletion. Recorded operation values are 203 ms to FTP,
203 ms from FTP, 156 ms replacement, and 63 ms deletion; the reported provider
watchdog setting is 8,000 ms. These observations are regression evidence, not an
improvement or a measured universal watchdog bound.

This host/bridge workflow complements, but cannot replace, the native Delete guard
fixture or the remaining native route/walker matrix. No live endpoint is exercised.
The runner disk audit reports 31 retained older-run issues; no older data was removed.
