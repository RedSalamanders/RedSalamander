# Operation Startrail Fresh/Resume closeout evidence

Status: **PASSED**

This curated bundle closes the final Operation Startrail validation campaign.
It retains the immutable shared validation plan, both plan decisions, both
terminal run states, and aggregate v2 records needed to review Fresh execution
and exact Resume provenance without checking in the reusable evidence cache.

## Fresh Full

- Run: `20260814T180639Z-35264-794aa0a44be34b5e987a64c945fb4f72`
- Mode/gate: Fresh / Full
- Verdicts: repository `PASSED`, change `PASSED`
- Coverage: complete; 18/18 entries executed and promoted
- Cases: 1,256 total; 1,203 passed; 0 failed; 53 documented skips
- Duration: 3,169,501 ms
- Commands broad seal: fresh
- Build receipt: `35845714ba25096fb3050de8b5d9f2a98768e662f7b2dcf2181d35ba000b4163`
- Snapshot: `f87a3d1e613ab1a23b09db0254d29e4157ac611b58e814f568e4dd1c281b8499`
- Plan: `6c9fad653c31a1eda8a1826c6f3f678a77b306fcb115aaf62cbe07c3340e4fc5`

The Fresh TestSandbox directory was removed by the exact Resume preflight's
bounded stale-run cleanup after terminal evidence had been published. Therefore
`fresh-aggregate.v2.json` was regenerated with the production
`New-RSValidationRunSummaryV2` builder from the retained immutable run state,
plan, decision, promotions, and entry evidence, using the terminal totals
captured from the successful Fresh invocation above. It is not represented as a
byte-for-byte copy of the removed TestSandbox file.

## Exact Resume

- Run: `20260814T190349Z-42220-48839b4b33244dc3b936af7e8c176ac8`
- Parent: `20260814T180639Z-35264-794aa0a44be34b5e987a64c945fb4f72`
- Mode/gate: Resume / Full
- Verdicts: repository `PASSED`, change `PASSED`
- Coverage: complete; 13 entries reused exactly, 5 executed, 18 promoted
- Executed: `standalone.filesystem-curl`, `perf.monitor-etw-latency`,
  `perf.performance-tests2`, `tooling.pester`,
  `tooling.vcpkg-merge-synthetic`
- Commands broad seal: exact-reused; no broad Commands process launched
- Duration: 142,312 ms
- Estimated avoided duration: 2,994,474 ms (49m 54.474s)
- Build receipt, plan digest, and workspace snapshot are identical to Fresh.

No source or build-artifact edit occurred between Fresh and Resume. The exact
Resume invocation used `-SkipBuild` and accepted the same content-addressed
receipt.
