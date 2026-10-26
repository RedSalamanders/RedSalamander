# File Operations P2.4 Debug integration gate

- Runner id: `20260823T213305Z-3636-896d5017b436403cb6dc4d9d964b2149`
- Command: `./Tools/Run-AllTests.ps1 -Suite FileOps -SkipBuild -ValidationMode Fresh`
- Debug x64 build attestation: `b7cf0f5a45f4302697eaea0bcf6bf4523ae7f66bf6f3360d9e2eeed6ff6a482d`
- Result: 112 passed, 0 failed, 27 documented environment/deferred skips; zero flaky,
  regression, isolation-suspect, or unclassified-failure classifications.

The run covers the bounded host bridge, Local recursive Copy through both single-item APIs,
long-path traversal, newly-created directory metadata restoration, pre-existing folder-merge
metadata preservation, cancellation/conflict routing, and the existing File Operations matrix.

`perf/perf_metrics.jsonl` is a deterministic original-order compaction of the P2.4 traversal,
deep-tree, and high-water metric families. The uncompressed local stream had SHA-256
`d3e6f4f3a62ad285f05eae7136e8ddbb6808900a7a47dfb04132a5a154380039`, 333,890,527 bytes,
and 914,939 rows. The retained compact stream has SHA-256
`b95c9c1bf873c15c8f617192b8667a28089bd5ac606a0e2aae0c2596004877fa`, 362,479 bytes, and
953 rows.
