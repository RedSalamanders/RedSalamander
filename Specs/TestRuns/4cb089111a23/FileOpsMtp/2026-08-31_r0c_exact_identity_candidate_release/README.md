# R0c MTP exact recovery identity candidate

- Machine profile: `4cb089111a23`
- Source commit: `00b6df42` (R0c exact recovery identity implementation and warning cleanup)
- Build: test-enabled x64 Release; full-solution receipt `ecc1c16341cd4cb6902bb3afaa722b72dada7faed9f2134fb06feb7309bfe9ca`; FileSystemMtp receipt `a5a3aa8988f6e1ef13f7c386f27e1863d27e093a39ffe5a7b45b2e010c127b1a`
- Command: `RedSalamander.exe --compare-selftest --selftest-case=mtp_r0c_exact_recovery_identity --selftest-fail-fast --selftest-no-activate --perf`
- Scenario: five sequential independent processes over 4,096 case-distinct full identities and one exact-PUID recovery record
- `FileOps.Mtp.R0c.ExactIdentity.SelfTestUs`: `25,939`, `27,596`, `26,463`, `26,550`, `26,466` us
- Nearest-rank metric p95: `27,596` us
- Whole-case durations: `40`, `39`, `37`, `39`, `39` ms; nearest-rank p95 `40` ms
- Resource shape: one 4,096-item backend enumeration vector plus constant match state; every sample reports `value0=4096` and `value1=4096`

Every focused contract is GREEN (`passed=1, failed=0`). The comparable whole-case p95 is 15 ms above the RED baseline and passes the safety ceiling `max(1.50 * 25 ms, 25 ms + 50 ms) = 75 ms`; every candidate metric is below 2,000,000 us. This is deterministic identity/recovery resource evidence and does not claim live-WPD latency.
