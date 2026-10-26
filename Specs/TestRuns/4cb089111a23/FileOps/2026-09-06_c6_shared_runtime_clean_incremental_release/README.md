# C6 shared-runtime cleanup — incremental Release and smoke

Ordinary test-enabled x64 Release build after the full rebuild: zero warnings/errors,
160,281 ms. Receipt:
`a6fece6675db82bf539822d2caf671a416ad4918b27e0c406522c637dff2576d`.
Source snapshot: `3754f54aea44516549094354eff5053a2588e68eb2fc6d939f6fec34518b7335`.
Some targets recompile/relink, so no no-op performance or universal link-skipping
claim is made.

Governed FileOps `R0fCurl_FakeFtpReadWriteCreateDelete`: 3 passed, 0 failed,
0 skipped (setup, CRUD, cleanup). Run:
`c6-runtime-clean-smoke-incremental-release-20260906`. This is loopback runtime/staging
smoke, not live remote mutation proof or Fresh Full. The runner's 185 disk-audit
issues remain visible in its preserved aggregate.

Raw copies, ordered metric selection and build-log normalization use the recorded
digest contracts. See the [repair evidence](../2026-09-06_c6_shared_runtime_clean/README.md)
for source copies, full receipts and cleanup-history snapshots.
