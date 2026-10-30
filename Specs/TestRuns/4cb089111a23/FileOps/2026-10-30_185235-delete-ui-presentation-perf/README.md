# Delete throughput and decision UI qualification

Protected scenarios: recursive permanent-delete completed-item sampling, bounded item/sec history, read-only consent repetition by typed file/folder scope, distinct access-denied routing, and standard terminal-button order with real keyboard navigation.

- Safety family: 8 passed, 0 failed, 0 skipped (183910 archive). This includes second-file-first admission and one read-only file consent covering both selected files; folders retain a separate scope.
- Rate display: 3 passed, 0 failed (184422 archive).
- Integrated presentation performance: 1 passed, 0 failed, 0 skipped, this archive.
- Live dialog/keyboard checks: 12 passed, 0 failed, 0 skipped (Commands/185637 archive).
- Native Debug builds: zero warnings and errors. Source contracts: 212 passed. Changelog and archive validators passed.

The safety/rate source snapshot precedes only the Commands mirror Tab expectation correction; runtime FileOps sources are identical. Each archive records its own source identity and build receipt rather than claiming identical complete source snapshots.

Performance: 954 popup renders, median 1645 us, nearest-rank p95 3061 us. Five graph-open batches of eight redraws had median 65966 us versus 65918 us for collapsed batches in the same run. These are Debug instrumentation observations, not a before/after release speedup claim. Existing deterministic integrated resource/churn checks passed. No clean pre-change baseline, Release, or Full-suite qualification is claimed.

Failed prior checks are retained in adjacent rejected archives. One failure assumed held Right would move away from now-rightmost Cancel; the expectation was updated to Left then Right. Another rejected whichever read-only sibling reached the prompt first; the fixture now accepts either sibling and rejects an extra same-scope prompt. Raw streams and build-invalidated runs remain intact in the marked sandbox. Curated metric streams preserve original row bytes/order; raw-summary.json binds them to complete retained streams by SHA256.

Application captures are actual deterministic harness output with invented fixture paths, never manual desktop captures. Interaction tests used the governed desktop lease and warning under the user's explicit authorization.
