# ViewerSpace x64 Release ABBA comparison

Campaign 20261009T043820114Z-bb8f9b16 completed at 2026-10-09T05:00:54Z. All 14 runner processes passed both selected cases (28/28); the two warmups are excluded from analysis. The 12 measurement observations form three serial ABBA blocks, six per side. Each of the six paired contrasts is retained in comparison-analysis.json, and the 14 per-run archives are adjacent under Viewers/ using the campaign labels.

Candidate receipt 8954e5665904ae14116dcf42b2e459de0718c25140827489d70a1fab0b8e3928 pins source snapshot 98a5ee498c9fd307fd6008f06b0d9fe73f40f413c8d93e6e53ddaeddad888745. Baseline receipt a56ada4920ea0ef1b69baaafd373e274e40177f0740d9ed5ff4ab694fe94e3b2 pins source snapshot e67a99083439ec38f6b54df71a50ac578fdf03339fca21310d69a90c89ef3091. The campaign manifest enforces those receipt identities for each run.

## Observed measurements

The analyzer reports descriptive paired ratios (candidate / baseline); values below 1 are lower observed measurements and do not establish a performance gain. Scan-total median ratios were 1.1832 for the small case and 1.4694 for the large case. With six pairs, nearest-rank p95 is the maximum observation: 5.9882 small and 1.6717 large, so it is not a stable tail estimate. The candidate scan medians were higher in this sample; the measurements do not establish cause or a general regression.

Candidate-only first meaningful paint missed the 250 ms target in 0/6 small and 0/6 large observations. Small median/p95 was 508.271/699.377 ms; large was 462.668/516.633 ms. The baseline did not measure this endpoint. Historical scan-start-to-first-paint timing is a separate metric, retains HRESULTs, and is not a substitute for the candidate meaningful-paint endpoint.

Memory is reported separately from scan timing. Median candidate/baseline peak-while-open private-byte-delta ratios were 1.0832 small and 1.0281 large; closed-after-quiet ratios were 1.1221 and 1.1079. Six pairs per scenario are too few to characterize tail behavior; see the complete per-phase distributions in comparison-analysis.json.

Both branches use matched queue diagnostics: producer-side aggregate counters with cumulative observer snapshots. The baseline observer/harness patch in this archive is copied from the prior campaign archive ../20261009T033043683Z-337faa08/baseline-observer-harness.patch, which is associated with the same baseline receipt/source-snapshot IDs. The current campaign manifest does not pin that patch hash directly; do not treat the inherited patch copy as a campaign-recorded patch identity. campaign-manifest.json does pin the three exact helper snapshots in reproduction_inputs; the corresponding preserved source copies here are .source.txt files and match the manifest SHA-256/byte lengths.

The raw campaign retained all 28 stdout/stderr logs. The campaign manifest records all 14 individual run IDs, receipt IDs, snapshot identities and outcomes. All run summaries record three sandbox-audit issues, so the successful selected cases are not a clean-sandbox or full-repository pass. No overall gain, causal source attribution, post-snapshot source qualification, or app ARM64 qualification is claimed.

manifest.txt binds each copied payload to its SHA-256, exact byte length and original source path; every pair was checked before/after copy. Archive-local .gitattributes is curator metadata excluded from the payload manifest. Per-run test data was not recopied here.