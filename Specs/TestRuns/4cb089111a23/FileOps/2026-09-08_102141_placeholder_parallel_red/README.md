# 2026-09-08_102141_placeholder_parallel_red

Governed run: `20260908T102141Z-27092-edc68532832c44c3bb0b344a664026ad`, x64 Debug, machine `4cb089111a23`.

Real product regression before the parallel-placeholder fix: native metadata exposes Cloud tag 0x9000601A/attributes 0x420, yet the parallel producer calls CopyFile after releasing its primary buffer. The placeholder contributes zero copied bytes and zero worker admissions. Only the adjacent junction is reported skipped. The fixture correctly fails the payload assertion.

The aggregate retains the build receipt and exact invocation. Per-case results and traces are unchanged. Metric compaction retains the named correctness/latency witnesses in original order with original and selected digests/counts; it is not a general throughput comparison.
