# 2026-09-08_095536_wof_fixture_diagnostic

Governed run: `20260908T095536Z-78560-571c660c454449318deca63e8acb6aa9`, x64 Debug, machine `4cb089111a23`.

NTFS D: was selected and FSCTL_SET_EXTERNAL_BACKING returned success, but no WOF reparse attribute was present. The case correctly reports skipped (wof-attribute-missing), so this is diagnostic evidence and does not close BR-4.

The aggregate retains the build receipt and exact invocation. Per-case results and traces are unchanged. Metric compaction retains the named correctness/latency witnesses in original order with original and selected digests/counts; it is not a general throughput comparison.
