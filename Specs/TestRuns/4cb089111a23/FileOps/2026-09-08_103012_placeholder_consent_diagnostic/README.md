# 2026-09-08_103012_placeholder_consent_diagnostic

Governed run: `20260908T103012Z-57260-b4d01a98856d4a22bd8d8b1f03bbcfa9`, x64 Debug, machine `4cb089111a23`.

The product fix now copies all 65,536 bytes with exactly one worker admission and skips the junction. The run remains FAILED because the test incorrectly expected hydration consent for an already fully hydrated file. Provider metadata reports basic/security features (0x11), without recall/placeholder-hydration risk. The final test must require no prompt for this fixture.

The aggregate retains the build receipt and exact invocation. Per-case results and traces are unchanged. Metric compaction retains the named correctness/latency witnesses in original order with original and selected digests/counts; it is not a general throughput comparison.
