# R0c MTP exact recovery identity baseline

- Machine profile: `4cb089111a23`
- Source commit: `6eb05929` (RED harness, before the R0c behavior change)
- Build: test-enabled x64 Release; full-solution receipt `cacc904cec570725a38fb18fa223605a9732baca645764e863500563d7db35fd`
- Command: `RedSalamander.exe --compare-selftest --selftest-case=mtp_r0c_exact_recovery_identity --selftest-fail-fast --selftest-no-activate --perf`
- Scenario: five sequential independent processes over 4,096 case-distinct fake MTP identities and one recovery record
- Whole-case durations: `25`, `24`, `23`, `25`, `23` ms
- Nearest-rank whole-case p95: `25` ms

Every baseline process is intentionally RED (`passed=0, failed=1`): the legacy case-folded identity hash collapses `CasePuid` and `casepuid`, so only 4,095 entries are exposed. That failure occurs before the recovery setup and therefore before `FileOps.Mtp.R0c.ExactIdentity.SelfTestUs` can be emitted. The archived whole-case duration is the honest pre-fix comparison; no synthetic baseline metric is substituted. The build contains the one test-hook C4189 warning later removed by `00b6df42`; it has zero errors.
