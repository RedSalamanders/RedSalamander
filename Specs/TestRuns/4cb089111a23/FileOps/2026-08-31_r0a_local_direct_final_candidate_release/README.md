# R0a Local direct-final candidate

- Machine profile: `4cb089111a23`
- Build: final test-enabled x64 Release, receipt `f431e72785d19f28da1dc0da92a3cdf341b874b95c6d56f36dda0970d159fcaf`
- Variant: exclusive final-leaf create with retained exact write/abort authority
- Scenario: five independent processes running `Floodgate_LocalCopyNewNameConcurrentReplacementSurvives`
- Metric: `FileOps.SelfTest.R0aLocalDirectFinalCopyUs`, 64 MiB ordinary absent-destination Local regular-file Copy
- Samples: `40,353`, `43,669`, `43,759`, `40,882`, `45,083` us
- Nearest-rank p95: `45,083` us
- Admission budget: `max(1.25 * 28,374, 28,374 + 25,000) = 53,374` us
- Result: PASS (`45,083 <= 53,374`); every sample is below the `2,000,000` us absolute cap

All five race runs passed `3/0/0`, with no classification, quarantine activity, timeout, or disk-audit issue. `abort/` separately proves Release abort failure returns retained-incomplete truth (`3/0/0`, disk audit `0`).
