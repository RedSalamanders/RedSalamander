# C6 shared-runtime cleanup — Debug rebuild and smoke

Full test-enabled x64 Debug rebuild: zero warnings/errors, 189,655 ms.
Receipt: `b69a108f2dd328eb9f134feabe2b2f868f40706f2849593253b96910ff2c5a2f`.
Source snapshot: `3754f54aea44516549094354eff5053a2588e68eb2fc6d939f6fec34518b7335`.

Governed FileOps `R0fCurl_FakeFtpReadWriteCreateDelete`: 3 passed, 0 failed,
0 skipped (setup, CRUD, cleanup), native suite 2,516 ms. Run:
`c6-runtime-clean-smoke-debug-20260906`. This is loopback runtime/staging smoke,
not live remote mutation proof or Fresh Full.

Raw results, trace, runner aggregate and entry fingerprints are copied exactly.
Metric selection preserves the declared original line order and raw-stream digest;
no throughput improvement is claimed. The build log's declared whitespace-only
normalization preserves its diagnostic content. The runner's 184 disk-audit issues
remain visible in its aggregate.

See the [repair and regression evidence](../2026-09-06_c6_shared_runtime_clean/README.md)
for failed baselines, source copies, full receipts and cleanup-history inspection.
