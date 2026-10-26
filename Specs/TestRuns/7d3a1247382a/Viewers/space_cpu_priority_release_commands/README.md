# CPU-priority Release Commands evidence

This archive records the x64 Release Commands run `20261009T060400Z-27392-e27b516b55d94b15a24a36aa38651812`: 31 passed, 0 failed, 0 skipped. The matching immutable build receipt and launch log are exact copies; their byte counts and SHA-256 values are in `build-provenance.txt`. The receipt pins source commit `c9dd2be6fcb6152d3cef9e2ae7ac91eca76d37c3` and source snapshot `17c73d11118294b1ce91d97c62a1d5cb36dd4652a9111c477fdcc1c9a24833e8`; build was x64 Release with tests enabled and exit code 0.

The compact performance payload consists of `perf_metrics.jsonl` followed by `perf_metrics-part2.jsonl`. Each shard matches its SHA-256, byte count, and row count in `evidence.json`; concatenation matches the declared retained-payload SHA-256. The curated evidence reports 5,586 retained rows selected from a declared 86,681-row source. The original run directory had already been removed by stale-run cleanup, so this archive-only check could not recompute the declared full-source hash or independently replay the source-line selection map.

Observed aggregate scan fields in `evidence.json` include 67 `viewer.space.scan.cpu_priority` records (value0/value1 are 8 in the aggregate summary), 67 `viewer.space.scan.first_preview_priority_jobs` records (value0 p50=1, p95=8, max=8), and 67 scan-total records. These are preserved observations, not a separate causal or performance-gain claim.

Capture images and original `captures/raw-provenance.json` are retained unchanged. Root directly reviewed and accepted all 11 whole frames at 144 DPI; see `captures/whole-frame-review.txt`. This applies only to these exact frames and scenarios. `* -text` keeps archived payload bytes from Git newline conversion.
