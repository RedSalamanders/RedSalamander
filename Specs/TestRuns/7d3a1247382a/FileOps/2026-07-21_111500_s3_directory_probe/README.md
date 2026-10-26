# S3 Interactive Directory Destination-Probe Evidence

Five test-enabled Release executions of `PluginContractTests.exe` passed against one rebuilt binary. Each execution runs the 16-object legacy and candidate paths in the same process with identical fake S3 graphs and a fixed 3 ms delay per metadata request.

The legacy path performed 32 destination refreshes and 64 underlying summary probes per repeat. The candidate performed 16 refreshes and 32 probes. Median measured probe time fell from 1,011,056 us to 504,449 us (50.11%). Correctness coverage in every execution includes byte preservation, no-callback fail-closed behavior, pre-authorized overwrite, Retry re-probe, and two stacked ancestor removals.

This is deterministic request-count and fixed-latency evidence, not a live S3 throughput claim. The JSONL commit field identifies the committed base; the implementation and evidence were captured from the documented working tree.
