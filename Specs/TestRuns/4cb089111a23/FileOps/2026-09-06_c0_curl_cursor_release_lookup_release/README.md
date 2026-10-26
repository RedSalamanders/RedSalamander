# Bounded cursor release: Release lookup

Final warning-corrected candidate after `96c5c5f8`. Full test-enabled Release
build: 248,705 ms, zero warnings/errors. Receipt and source snapshot are in
`build-provenance.json`; `tested.patch` and `tested-source.json` preserve the
seven-file source delta before closeout documentation. The first warning-bearing
build/run is retained in `../2026-09-06_c0_curl_cursor_release_initial_lookup_release/`.

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild
-TestRoot C:\RedSalamander.Perf -CaseFilter C0_CurlEntryLookupTruth
-RunId c0-curl-cursor-release-final-lookup-release-20260906`.

Result: 3/0/0 passed/failed/skipped, 24.1 seconds; native lookup 284/284.
Eleven added assertions cover bounded read-ahead, completed parent/child transport
reuse, unread parent records, empty and 4,097-row listings, abandonment,
cancellation, bad final reply and healthy reborrow. The ten lifetime metric rows
all pass; nested small listings use one authenticated session for two LISTs.
The existing fragmented-row and fail-closed exact-lookup checks remain enabled.
Final source/resource policy: 197/197, zero skips, 18.1170054 seconds.

Exact copied evidence and ordered metric selection are independently digest-bound
by `file-provenance.json` and `metric-selection.json`. The transformed build log
records both raw and archived digests. These are focused results, not Fresh Full
or proof of real FTP/SSH quotas or whole-operation memory bounds.
