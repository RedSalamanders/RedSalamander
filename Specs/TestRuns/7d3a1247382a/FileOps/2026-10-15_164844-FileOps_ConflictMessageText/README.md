# FileOps_ConflictMessageText: focused Debug regression

Result: **3 passed, 0 failed, 0 skipped**; runner exit **0**. The three recorded cases are Setup, FileOps_ConflictMessageText, and Cleanup_RestorePluginConfig. Raw per-case status, trace and result files are preserved unchanged.

Run `20261015T144816Z-49600-87c5b8ef6c864d4abb18beeb9831b7bc` uses test-enabled **x64 Debug**, build receipt `047f98907061f915d00f1f490b967e89fab62bf15c3f8aa6169c531d0240b5af`. This is focused correctness/observational evidence; it does not establish Full-repository or Release qualification or a performance improvement.

The runner reports **3 legacy sandbox audit findings** (`unexpected-test-run-dir`): `archive-skip-check-7756799779d4`, `dxui-ci-first-37835055279`, and `dxui-ci-retry-37835055279-attempt2`. Thus the scenario passed while the sandbox audit was not clean; no unrelated directories were removed.

Metrics remain the complete original 101,825-byte stream, SHA-256 `afc085814d1fd82ef6ef6d0d0cfd9bba2fbd34e3db47c8a02c98f10ac2e5a375`; the external copy remains under `D:\RedSalamander.Perf\evidence\fileops-ux-review-20261015\20261015T144816Z-49600-87c5b8ef6c864d4abb18beeb9831b7bc\selftest\last_run\perf\perf_metrics.jsonl`.
