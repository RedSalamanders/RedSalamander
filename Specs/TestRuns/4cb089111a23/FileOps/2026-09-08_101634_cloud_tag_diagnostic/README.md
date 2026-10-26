# 2026-09-08_101634_cloud_tag_diagnostic

Governed run: `20260908T101634Z-95936-f4fd60d94cdf4c3ea3a87b6a54fa3355`, x64 Debug, machine `4cb089111a23`.

Cloud conversion succeeds, but this selftest process observes path attributes 0x20, handle attributes 0x20 and tag 0. Both native queries succeed. The same fixture in PowerShell exposes attributes 0x420/tag 0x9000601A. This isolates a process/thread placeholder visibility difference; BR-4 remains open.

The aggregate retains the build receipt and exact invocation. Per-case results and traces are unchanged. Metric compaction retains the named correctness/latency witnesses in original order with original and selected digests/counts; it is not a general throughput comparison.
