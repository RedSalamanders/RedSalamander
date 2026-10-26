# Cinderstar bounded WSL catalog validation

- Date: 2026-08-09
- Machine profile: `4cb089111a23`
- Source fingerprint: `e0de6582ee82eaa5bb0711260d04d5b6642a9660` plus the uncommitted Cinderstar Package 4 implementation
- Build: x64 Debug
- Focused entry point: `PluginContractTests.exe --registry-wsl-catalog`
- Result: 25 passed / 0 failed
- Metric: `wsl.catalog.enumerate_us` (`durationUs`; one row per complete catalog call)
- Measurement: 136 us for five inspected volatile subkeys and two retained records
- Build proof: `.build/logs/msbuild-20260809_121335_273.log` (0 warnings, 0 errors)
- Full solution rebuild: `.build/logs/msbuild-20260809_120926_225.log` (0 warnings, 0 errors)

## Performance gate

- Scenario: opening a navigation menu that discovers registered WSL distributions.
- Change type: correctness and responsiveness stabilization.
- Risk protected: malformed or changing registry values causing out-of-bounds reads, unbounded allocation/retry,
  catalog disagreement, or synchronous menu latency.
- Instrumentation: `wsl.catalog.enumerate_us`, with retained-record count in `value0` and inspected-subkey count
  in `value1`.
- Deterministic validation: a volatile per-user registry fixture covering terminated/nonterminated, long, empty,
  odd-byte, wrong-type, over-cap, and synchronously mutated values plus filtering, sorting, base paths, and WSL1/2.
- Bounds: 64 KiB default string cap (1 MiB hard cap), three default read attempts (eight hard maximum), and 1,024
  distribution subkeys.
- Expected direction: bounded menu discovery independent of malformed value length or endless key enumeration.
- Evidence quality: single-sample Debug diagnostic baseline; no percentile, throughput improvement, or Release-latency
  claim is made. The row exists to establish the metric and deterministic baseline for later comparisons.
- Analyzer: `Tools/Show-PerfRuns.ps1 -RunsRoot Specs/TestRuns -Scenario wsl_catalog_volatile_registry
  -Metric wsl.catalog.enumerate_us -ShowBuildFlavor`; sample-quality gating is not applicable to this
  single-sample non-percentile baseline.
