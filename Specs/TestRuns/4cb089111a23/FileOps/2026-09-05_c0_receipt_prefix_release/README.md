# C0 bounded mutation receipt — test-enabled Release

Run `c0-receipt-prefix-enabled-release-20260905`, `C0_`: 6/6 passed, no skips
(four behavior cases plus setup/cleanup). Build receipt
`b1cb2ab29e8ca3c34ccb3f3205d05444ebdd11eef1ee502519f04a073f0c9bc4`;
`tests_enabled=true`, zero warnings/errors. Base `58f22d0e` plus working changes
is runner-bound. The profile-root SQLite runtime is present in this accepted build.

The same four behavior cases as `2026-09-05_c0_receipt_prefix_debug` pass under
optimization: exact V1/header guard-page bounds; malformed and missing receipts
cannot inherit old proof; future suffixes and stage validation; ignored malformed
callback error still produces uncertainty; one request/commit for lost-response
Copy/Delete; independent Local publication/cleanup truth with one ordinary prompt.

The prefix-boundary metric is 90 us with five rejected header shapes and five host
callbacks. Debug measured 249 us on the same machine, but configurations differ;
these are correctness witnesses, not a performance improvement/percentile comparison.
All raw result/trace/aggregate/metric files are retained unchanged.

The prior non-test Release invocation failed to collect a receipt because the root
SQLite DLL was absent. Its separate `2026-09-05_c0_release_staging_failure` archive
remains unresolved by this successful test-enabled profile. The disk audit here flags
27 retained earlier run directories, not deleted or claimed clean. Complete route
coverage, remaining C rows and final Fresh Full are still required.
