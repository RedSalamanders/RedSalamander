# DiscoveryScope_MultipleFoldersLiveProgress: focused Debug regression

Result: **3 passed, 0 failed, 0 skipped**; runner exit **0**. The three recorded cases are Setup, DiscoveryScope_MultipleFoldersLiveProgress, and Cleanup_RestorePluginConfig. Raw per-case status, trace and result files are preserved unchanged.

Run `20261015T145057Z-49600-a2a3cf1a5ca7412c89ba7bf31001fa3f` uses test-enabled **x64 Debug**, build receipt `047f98907061f915d00f1f490b967e89fab62bf15c3f8aa6169c531d0240b5af`. This is focused correctness/observational evidence; it does not establish Full-repository or Release qualification or a performance improvement.

The runner reports **3 legacy sandbox audit findings** (`unexpected-test-run-dir`): `archive-skip-check-7756799779d4`, `dxui-ci-first-37835055279`, and `dxui-ci-retry-37835055279-attempt2`. Thus the scenario passed while the sandbox audit was not clean; no unrelated directories were removed.

Metrics remain the complete original 481,115-byte stream, SHA-256 `fa8d7dd74f16287ed7c904d0b2f93261bcf92da5575e6cc16e0f1d73e10f54cc`; the external copy remains under `D:\RedSalamander.Perf\evidence\fileops-ux-review-20261015\20261015T145057Z-49600-a2a3cf1a5ca7412c89ba7bf31001fa3f\selftest\last_run\perf\perf_metrics.jsonl`.
