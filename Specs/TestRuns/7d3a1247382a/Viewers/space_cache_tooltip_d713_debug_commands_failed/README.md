# ViewerSpace d713 Debug Commands failed run

Run `20261009T090204Z-13052-c39359d4a7204a038bac43bb0e83edd9` used x64 Debug Commands with case filter `cmd_viewer_space_`, 33 cases total: 30 passed, 1 failed, and 2 skipped. The two opt-in cases were skipped, not passed.

The only failed case was `cmd_viewer_space_hover_does_not_rebuild_static_cache`. It failed on sample 11: the expected stable-hover tooltip paint had not occurred before containment. The exact failure details and attempt are preserved in `run-all-tests-results.json` and the case history. The run summary classifies it as `UNCLASSIFIED_FAILURE`.

The run summary also records an unclean sandbox audit with three unexpected run-directory findings. This is failed diagnostic evidence, not a clean Commands pass or overall qualification.

Build receipt ID: `a1de121d1b5ad39f29ff3e7e2f8de43e041472467c9ed8551c67d1c8f24600f0`. Source snapshot: `b1280fd000244b10dad4513f6605442a01b1d25c25beca54fec7a74f79782d90`. The exact receipt and launcher transcript are preserved in this directory. Receipt: 95,452 bytes, SHA-256 `c0cdbc5244e2d3031169a226c9b7689ccec2cecf020c3d09cd4afc81b791777c`. Launcher: 3,543 bytes, SHA-256 `2b2e29d452f65f9933e0ed531d0ad7c783a16fe004e41315a497023c48efb98c`.

The primary archive retains five scenario PNGs. The companion `_boundary_captures` archive retains six threshold PNGs. Root has inspected only `boundary-width-503.png` for diagnosis; that inspection is not publication acceptance. Whole-frame publication review has not been completed for either archive.

Full perf source: `D:/RedSalamander.Perf/runs/20261009T090204Z-13052-c39359d4a7204a038bac43bb0e83edd9/artifacts/selftest/last_run/perf/perf_metrics.jsonl` (23,376,768 bytes, SHA-256 `b7892e1fd672f7ddbf8fbe6002758ffa5e232acb4a90cd7a0ae306fccdc2225d`, 58,320 rows). The identical complete file is preserved at `D:/RedSalamander.Perf/evidence/viewer-space-final-qualification-d713/commands-debug-failed.perf_metrics.jsonl` with the same byte length and SHA-256.

The archived bounded `perf_metrics.jsonl` shard contains 5,067 rows (1,906,844 bytes, SHA-256 `735fef5cc20d9f948967c71f99d759b857ab3fd66d3e8ef5a19c5b2ea00d3b3b`). Its ordered source-line references in `evidence.json` were replayed against the original file; all retained lines matched exactly and the shard digest matched. This is a retained subset, not a replacement for the complete source.

Local `.gitattributes` sets `* -text` so archived payload bytes are not text-normalized; it is curator metadata.
