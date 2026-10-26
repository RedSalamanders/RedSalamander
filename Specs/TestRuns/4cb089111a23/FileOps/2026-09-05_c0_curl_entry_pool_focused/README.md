# Cursor cache retention — focused pass

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_CurlEntryLookupTruth -FailFast -RunId c0-curl-entry-pool-debug-20260905`.
Based on `3dab443c`; pending pool/cursor/test changes. Build 03:04, zero warnings/errors.
Receipt `7f21414498dceac23005ab6a1481dfb3d17e8328a507272e26a6132a62177b1e`;
workspace `e2d87c13c40c5dcf4008977aca6a9252cb595b8eedcf2459df5192e02846d2f8`.

197/197 assertions, 3/3 harness cases, zero skips. All eight canceled admissions
allow a subsequent same-endpoint lookup. Each 16-query wide probe uses one passive
listener with 15 reuses. Unix/DOS durations are 267,142/261,366 us, versus
386,833/350,564 us in the bounded pre-pool candidate on this machine. Owned path and
metadata reservations remain 208/51,400 bytes, excluding opaque transport state.
Requested metadata and all 4,097 rows remain checked.

Full perf: 148,388 bytes, 402 rows, SHA-256
`e5c592241f6310be1496a29bf384862777d8f9c0687573d29cf0d7975affff2b`.
Case 6,922 ms, suite 8,281 ms. Disk audit 80 issues; none deleted.
The [same-source mixed-workload failure](../2026-09-05_c0_curl_entry_pool_mixed_red/README.md)
prevents qualification. A pure lookup pass does not establish transfer/cache reuse.
