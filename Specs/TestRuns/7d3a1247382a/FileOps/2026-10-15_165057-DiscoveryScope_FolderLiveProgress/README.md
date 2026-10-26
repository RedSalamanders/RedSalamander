# DiscoveryScope_FolderLiveProgress: focused Debug regression

Result: **3 passed, 0 failed, 0 skipped**; runner exit **0**. The three recorded cases are Setup, DiscoveryScope_FolderLiveProgress, and Cleanup_RestorePluginConfig. Raw per-case status, trace and result files are preserved unchanged.

Run `20261015T144909Z-49600-80ef7a469b39482097c865f371dcee79` uses test-enabled **x64 Debug**, build receipt `047f98907061f915d00f1f490b967e89fab62bf15c3f8aa6169c531d0240b5af`. This is focused correctness/observational evidence; it does not establish Full-repository or Release qualification or a performance improvement.

The runner reports **3 legacy sandbox audit findings** (`unexpected-test-run-dir`): `archive-skip-check-7756799779d4`, `dxui-ci-first-37835055279`, and `dxui-ci-retry-37835055279-attempt2`. Thus the scenario passed while the sandbox audit was not clean; no unrelated directories were removed.

Metrics are an ordered byte-exact FileOps selection in three JSONL parts; see [selection manifest](metrics-selection-manifest.json) for the raw SHA-256, row indices and ordered part digests. The complete raw stream remains at the external path recorded there. No percentile or throughput-improvement conclusion is drawn.
