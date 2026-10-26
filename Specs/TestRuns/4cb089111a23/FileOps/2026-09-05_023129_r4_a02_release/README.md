# R4-A02-2 focused Release evidence

- Governed run: `20260905T003040Z-40676-32154703eff145d6a20f8a71b589249d`
- Build receipt: `7c133cac9ee36c799c98667b36c6badbc465f37bc1d42eabe90a5b502e014c1b`
- Result: 9 passed, 0 failed, 0 skipped; disk audit clean.
- Exact filter: `R4A02_` on test-enabled x64 Release.
- Claim: exact Run relations admit only disclosed tasks; the bounded live-output index reaches a measured high-water of 201 entries (below its 256-entry cap); the forced overflow takes the conservative guard path; Queue waits for the exact publisher; Skip and explicit invalidation retain their tested consequences; a third undisclosed task prompts independently.

The raw `perf_metrics.jsonl` remained in the governed test root because it was
16,427,948 bytes (46,788 rows), above the curated-run size limit. Its SHA-256 was
`c7a0a1a95d263effefd1d1cdda872b19c11dee53546565763b871af1393f6fa6`.
The compact JSONL beside this file preserves the claim-bearing source rows in their
original order. The retained source row numbers are 2307, 4600, 6739, 45077,
45100, 45553, 45576, 45577, 45721, 46011, and 46054.
