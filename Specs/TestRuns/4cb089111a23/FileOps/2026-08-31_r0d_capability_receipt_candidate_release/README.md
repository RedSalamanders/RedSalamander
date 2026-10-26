# R0d receipt and capability honesty candidate

- Machine profile: `4cb089111a23`
- Source commit: `57f5ab1a` (R0d implementation)
- Build: test-enabled x64 Release; full-solution receipt `0f1b7b85a4ab2d0a49b3ef35e318bfd50ae80f52cb035f511e967c793128da7a`; diagnostics `0 warning(s), 0 error(s)`
- Command: `RedSalamander.exe --fileops-selftest --selftest-case=FileOps_ProviderCapabilityMatrix --selftest-no-activate --selftest-timeout-multiplier=4`
- Scenario: five sequential independent processes, each running 4,096 receipt classifications, 8,192 host admission decisions, and 12,288 capability queries
- `FileOps.SelfTest.R0d.CapabilityReceiptHonestyUs`: `29,704`, `29,940`, `29,497`, `29,485`, `29,951` us
- Nearest-rank p95: `29,951` us
- Admission budget: `max(1.50 * 31,497, 31,497 + 50,000) = 81,497` us
- Result: PASS (`29,951 <= 81,497`); every sample is below the `2,000,000` us absolute cap
- Resource shape: one capability document and constant typed-result state; every sample reports 12,288 queries, 8,192 admission decisions, and 4,096 classifications

All five focused FileOps runs passed `3/0/0` with no skips. Exact-Release provider qualification also passed: `PluginContractTests.exe` returned 0, Microsoft Drive debug contracts reported `223/0`, Curl `226/0`, S3 multipart `50/0`, S3 directory probes `157/0`, and the focused MTP instance-honesty case passed `1/0/0`. This is in-process admission/classification evidence, not live-provider latency.
