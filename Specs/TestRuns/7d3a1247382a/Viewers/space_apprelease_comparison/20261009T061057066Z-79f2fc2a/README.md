# ViewerSpace x64 Release ABBA comparison

Campaign `20261009T061057066Z-79f2fc2a` completed 2026-10-09 06:32:50 UTC. All 14 runner processes exited successfully and passed both selected Commands cases (28/28). Two warmups are excluded; the 12 measurements form three serial ABBA blocks with six observations per side. The exact campaign manifest, 28 stdout/stderr logs, both build receipts, analysis output, and the three helper source snapshots pinned by the manifest are preserved here. `manifest.txt` binds each copied payload to its original source path, byte count, and SHA-256. `.gitattributes` is curator metadata outside the payload manifest. Per-run archives remain in adjacent campaign-labeled directories.

## Recorded identities

- Baseline receipt `a56ada4920ea0ef1b69baaafd373e274e40177f0740d9ed5ff4ab694fe94e3b2`, source snapshot `e67a99083439ec38f6b54df71a50ac578fdf03339fca21310d69a90c89ef3091`.
- Candidate receipt `039e7389819621e3195ffeb57013b76c98cd1c3b4050cd42645235ea3ef043e5`, source snapshot `17c73d11118294b1ce91d97c62a1d5cb36dd4652a9111c477fdcc1c9a24833e8`, source commit `c9dd2be6fcb6152d3cef9e2ae7ac91eca76d37c3`.
- Both sides used the matched queue-observer policy recorded in the campaign manifest. The manifest pins the three exact runner/curator/analyzer helper copies under `reproduction/`; no other baseline patch identity is inferred here.

## Analyzer observations

The pinned analyzer calls these descriptive measurements and asserts no positive performance gain. Candidate/baseline scan-total median ratios were 0.0593 small and 0.0511 large; nearest-rank p95 ratios were 0.2652 and 0.0636. With six pairs, nearest-rank p95 is the maximum observation and is not a stable tail estimate.

Candidate-only first meaningful paint has no baseline comparison. Against the 250 ms target, small measured 264.623 ms median / 282.865 ms p95 with 0/6 meeting target; large measured 164.917 / 194.376 ms with 6/6 meeting target.

Median candidate/baseline private-byte-delta ratios were 1.0519 small and 1.0340 large peak-while-open; closed-after-quiet they were 1.1060 and 1.1140. Closed-after-quiet working-set ratios were 1.0357 and 1.1206. These are separate memory observations and do not establish an overall gain.

The earlier campaign archive `../20261009T043820114Z-bb8f9b16/` used candidate receipt `8954e5665904ae14116dcf42b2e459de0718c25140827489d70a1fab0b8e3928` and earlier source snapshot `98a5ee498c9fd307fd6008f06b0d9fe73f40f413c8d93e6e53ddaeddad888745`; it reported scan-total median ratios 1.1832 small and 1.4694 large, with no established cause. This later campaign uses receipt `039e7389819621e3195ffeb57013b76c98cd1c3b4050cd42645235ea3ef043e5` and source snapshot `17c73d11118294b1ce91d97c62a1d5cb36dd4652a9111c477fdcc1c9a24833e8`. The earlier negative measurement remains preserved as historical evidence; this subsequent comparison does not establish cause or qualify the final test-only serial-hook source, whose qualification is still pending.
