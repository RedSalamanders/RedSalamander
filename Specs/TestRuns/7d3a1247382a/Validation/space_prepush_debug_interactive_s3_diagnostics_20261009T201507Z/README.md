# Debug PrePush diagnostics evidence

Run `20261009T201507Z-45224-aa600fc03f9e4418b3b3e7152722bb0f`. Original overall verdict: **FAILED**. This archive preserves scoped reports or diagnostic logs only; it is not a qualification pass. Original source paths and surviving preserved copies are separate manifest columns. Availability is true only when the original path still exists with identical bytes; unavailable sources are never reconstructed.

ViewerPETests.Interactive output includes failures in restored foreground/keyboard focus and stopping before another child. PluginContractTests output records three FileSystemS3 failures: recursive Delete did not list the fixture prefix; a request with a body still arriving did not return after Cancel; and Cancel did not return a streaming request within a few body callbacks. The logged 15,000 ms return is supporting timing evidence for the cancellation failures. The outer runner log is retained byte-exact. These are preserved diagnostics, not a new reproduction or a claim that every overall failure has been assigned a cause.

[Run reports and validation checkpoint](../space_prepush_debug_failure_reports_20261009T201507Z/README.md).
The 2 MiB/file and 5 MiB/archive limits are enforced before writing. The oversized 2,269,578-byte perf_metrics.jsonl is intentionally omitted; fixture/work trees are not selected.
