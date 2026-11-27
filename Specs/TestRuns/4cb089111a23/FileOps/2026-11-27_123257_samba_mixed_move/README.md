# I66 mixed Local and live SMB Move

The test-enabled x64 Release solution built in 10 minutes 1 second on `519dad224f73eec5f2d6bd03c80323df4306ad92` with zero warnings and errors. Verified receipt `78af1685c311d199f65fd88a177c7563e4fa8ad3dcb6cc50c603e3afbf7e3e66`. Static, pinned formatting and changelog checks passed.

Each run used `FileOps_SmbMixedMove` with the activation blocker. The source worktree, binaries and runner policy matched on both endpoints. Each run passed 3 checks with zero failures and skips. The runner preserved filtered repository status `NOT_EVALUATED`.

| Live source | SMB cleanup authority | Exact run | Task through cleanup (µs) |
| --- | --- | --- | ---: |
| RubyGloom Samba | Corroborated source facts | `20261127T113257Z-59436-093e8ff4480e4d279d7f4212f4e2c843` | 193,174 |
| Windows SMB | Exact source identity | `20261127T113410Z-69956-486fd730fd7345818cd14864ba116779` | 39,171 |

The Samba source opened with no exact identity and the provider returned matching source facts. The native `FileOps.Move.SourceRemovedByFacts` metric recorded one successful source removal. Windows SMB used exact binding. In each run a single Local Native and SMB Managed task published byte-identical destinations, removed both originals, used one Managed stage, and retired the fresh empty SMB fixture nonrecursively. The Windows control restored the resource manifest byte for byte.

Native result SHA-256 values are `3fd2f513bbcabc50a74a77d032210f8c22430ee376e514f78e143f29a9d8fdb7` for Samba and `37c3445c0444fca35f9f55d69dfe91e763e26f46244effce1a7b56a0999387b1` for Windows SMB. Both are under `artifacts/selftest/last_run/fileops` in the completed-run archives. The runner reported the same pre-existing 120 sandbox disk audit issues and retained nine older run roots.
