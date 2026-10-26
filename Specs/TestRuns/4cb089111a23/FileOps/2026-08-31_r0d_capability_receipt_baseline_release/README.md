# R0d receipt and capability honesty baseline

- Machine profile: `4cb089111a23`
- Source commit: `bf037d2d` (RED contracts, before the R0d production change)
- Build: test-enabled x64 Release; full-solution receipt `14c9932b3138cfc554de98fb660a8e0eed992679b729db854ffba81c291d81dd`
- Command: `RedSalamander.exe --fileops-selftest --selftest-case=FileOps_ProviderCapabilityMatrix --selftest-no-activate --selftest-timeout-multiplier=4`
- Scenario: five sequential independent processes, each running 4,096 receipt classifications, 8,192 host admission decisions, and 12,288 capability queries
- `FileOps.SelfTest.R0d.CapabilityReceiptHonestyUs`: `30,097`, `31,497`, `30,163`, `30,011`, `30,383` us
- Nearest-rank p95: `31,497` us
- Resource shape: one capability document and constant typed-result state; every sample reports 12,288 queries, 8,192 admission decisions, and 4,096 classifications

Every baseline process is intentionally RED (`passed=2, failed=1, skipped=0`): a successful receipt-less destructive Native result reached `Completed`. The protected deterministic loop and metric complete before that exact assertion; no synthetic baseline is substituted.
