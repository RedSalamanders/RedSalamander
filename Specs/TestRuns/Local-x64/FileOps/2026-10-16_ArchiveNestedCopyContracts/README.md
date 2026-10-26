# Nested archive path contract regression evidence

Scope: standalone plugin path contracts, x64 Debug, 2026-10-16. This is not a Full repository verdict or a throughput comparison.

The new PluginContractTests assertions fail against the pre-fix on-disk Debug archive provider: three mixed/backslash path derivations and the backslash child-name rejection fail. `baseline.txt` preserves the complete `PluginContractTests.exe --package-smoke` output (exit 1). This is a negative correctness control, not an attested baseline solution or performance baseline.

After rebuilding the candidate, the canonical affected runner's full PluginContractTests entry passes, including all seven new archive assertions and the existing typed/diagnostic contract consistency checks. `candidate.txt` is its complete original output. The immutable final candidate build receipt is retained beside it. The source base is 5141305d9c1d284c031d325d2e448fcd4a13a1d9 plus the working-tree fix and tests bound by that receipt.

Raw candidate run: E:/RedSalamander.Perf/runs/20261016T130548Z-6680-4e2711b647d54c1d8cf87c170b71a771/. The broader affected run is still in progress when this standalone evidence is promoted.

FileOps_ArchiveNestedCopyLocal and FileOps_ArchiveNestedCopyUnc are implemented and compiled. Their real Copy scenarios remain PENDING_INTERACTIVE until the user agrees to the governed desktop lease. The loopback UNC sandbox was reachable in a read-only preflight. No SMB Copy or per-copy performance measurement is claimed here. No original movie payload was copied, and no NAS or large multipart RAR qualification is claimed.