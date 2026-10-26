# IMAP structural attributes: Release candidate

Focused result: **208/208 assertions pass**, three runner cases passed and none
failed/skipped. The 12 standalone IMAP helper functions also pass after all 177
receipt artifacts are verified. [Broad C0 qualification](../2026-09-06_c0_imap_listing_attributes_broad_release/README.md)
passes 18/0/0 separately.

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild
-CaseFilter C0_CurlImapListingTruth -TestRoot C:/RedSalamander.Perf
-RunId c0-imap-attributes-candidate-release-20260906`.

The exact tested source is the
[Debug candidate](../2026-09-06_c0_imap_listing_attributes_candidate_debug/README.md)
patch and source-hash manifest against `f5858fbe`, with the same snapshot:
`57bc7d9f6a1ddc2e80539eec07b74433272555dedc3d9cc4d842d78a580c756c`.
Full test-enabled Release receipt:
`c38aa434016840c80c9f10fdc99067005f64010deb8e33b3883156bbf243e373`.
Its full build has zero warnings/errors and takes 249,813 ms.
The shared source/resource policy result is 197/197 in the Debug archive.

All baseline assertions remain, with the additional 74 controls documented in the
Debug comparison. Complete listings preserve request counts 1/2/37 and summary
peaks 1/200/200 for 1/201/4,097 messages. The parser and repair behavior is covered
through the actual synchronous production owner, not a second parser model.
No source/spec/index changes, retries, deadline increases or settings bypasses
occurred during the qualification sequence.

Preserve the selected metrics in original order and the raw failed baseline.
This is correctness and request/memory-shape evidence, not a throughput,
percentile, whole-buffer memory or live-IMAP/TLS claim. The runner retains 207
disk-audit findings. Final Fresh Full remains a separate gate.
