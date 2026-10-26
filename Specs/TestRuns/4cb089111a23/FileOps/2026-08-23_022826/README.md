# File Operations P1.5 Inline F2 Debug evidence

This governed x64 Debug run records the P1.5 Inline F2 worker, Queue-bypass,
identity-aware interlock, and cancellation baseline on machine `4cb089111a23`.

- Focused result: 3 passed / 0 failed / 0 skipped.
- Focused governed run id: `20260823T002807Z-26200-d5c7f0baff1e43068ef7f564fb7b5a38`.
- Complete File Operations aggregate: 112 passed / 0 failed / 27 skipped, run
  `20260823T000732Z-42220-f73014eafbbb49bdb51baf71cee20367`.
- Final build attestation: `fd95dad5bfc45aded7855ed9f7fc7eea53be3c047dc95e008092a2527d17806f`.
- Embedded source commit: `689c930ca4c1ab33c61b2f790d5667a47070f835`, plus the P1.5 working-tree
  implementation and specification diff archived by this commit.

The focused case proves that Inline F2 executes on a worker, bypasses the global
Queue admission wait, serializes overlapping parent mutations, completes both
successful renames, and cancels an interlock waiter before it reaches the
provider mutation boundary. The final source also compares retained object
identity across different roots in the same provider identity domain; the full
cross-root alias matrix remains part of the active P1.6 stop-condition gate.

No throughput improvement is claimed. The metrics are a safety-path baseline;
the deliberate self-test pause dominates the owner execution durations.

## Compact metric manifest

The original focused metric stream is checked in as `perf/perf_metrics.jsonl`.
`perf/perf_metrics_claims.jsonl` preserves the 13 P1.5 rows in their original
order, and `perf/perf_metrics_summary.json` binds the compact claim set to that
original stream.
