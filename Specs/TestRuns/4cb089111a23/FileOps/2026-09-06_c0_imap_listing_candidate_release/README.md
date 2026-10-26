# IMAP listing candidate — Release

Focused result: 70/70 assertions, runner 3 passed, zero failures/skips.
The 12 standalone IMAP helper functions pass after verification of all 177 artifacts.
Broad C0 passes 18/0/0 in the [companion archive](../2026-09-06_c0_imap_listing_broad_release/README.md).

Full test-enabled Release receipt:
`5e35f358b038381668b629398b7bc62b89dcde2934b7a8bbe750f55639893ab7`.
Build: 250,502 ms, zero warnings/errors. Snapshot
`fa008de406f84ba9fd25319d492403a4e60f8cc1059ca8e34daf520fcd2675b9`
matches the [Debug tested inputs](../2026-09-06_c0_imap_listing_candidate_debug/tested-source.json).
No source/test/tool change occurred between build and qualification.

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -CaseFilter C0_CurlImapListingTruth -TestRoot C:\RedSalamander.Perf -RunId c0-imap-listing-candidate-release-20260906
```

Helper invocation uses the exact function retained in the Debug archive, with
`-Configuration Release -RunId c0-imap-listing-helpers-release-20260906`.
The Debug archive describes assertions, comparison and remaining limitations.
This Release run has no matching pre-repair Release baseline and makes no speedup
claim. The 4,097-message summary peak is at most 200, without dropping messages.
The runner retained 200 historical disk-audit issues. No saved profile, TLS policy,
remote data or user-owned recovery files were changed. This is not Fresh Full proof.
