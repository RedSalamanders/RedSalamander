# C6 shared-runtime cleanup — Release rebuild and smoke

Full test-enabled x64 Release rebuild: zero warnings/errors, 257,128 ms.
Receipt: `45f63abadf35472b8da15a4dd0dc5743f30cc70f39822a45b40ca52f9b327e7f`.
Source snapshot: `3754f54aea44516549094354eff5053a2588e68eb2fc6d939f6fec34518b7335`.

Governed FileOps `R0fCurl_FakeFtpReadWriteCreateDelete`: 3 passed, 0 failed,
0 skipped (setup, CRUD, cleanup), native suite 1,906 ms. Run:
`c6-runtime-clean-smoke-release-20260906`. This is loopback runtime/staging smoke,
not live FTP/SFTP/SCP/S3/IMAP mutation proof or Fresh Full.

Raw results, trace, runner aggregate and entry fingerprints are copied exactly.
Metric selection preserves the declared original line order and raw-stream digest;
no throughput improvement is claimed. The build log's declared whitespace-only
normalization preserves its diagnostic content. The runner's 183 disk-audit issues
remain visible in its aggregate.

See the [repair and regression evidence](../2026-09-06_c6_shared_runtime_clean/README.md)
for failed baselines, source copies, full receipts and cleanup-history inspection.
