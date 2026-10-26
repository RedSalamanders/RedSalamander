# Rejected default-options native seed pool

Source `7ce7fcee0da65c4084c6f1fe1deb25f9d9b387d3` built x64 ASan Debug
and Release without warnings/errors. Both parser suites, all six native-MFT
memory/timing scenarios and seven ASan service regressions passed: 17 focused
native checks, zero failed/skipped checks. The service regressions included the
unchanged stale-root warmup watchdog and the original three service failures.
These focused passes do not clear the earlier failed Full.

The independent paired comparison rejected the candidate:

| Profile | Seed wall-time median ratio | Sampled private-memory median ratio | Verdict |
| --- | --- | --- | --- |
| ASan Debug | 0.623264 | 1.369510 | FAILED |
| Release | 0.908935 | 3.375519 | FAILED |

ASan sampled private memory rose from 3,199,954,944 to 4,382,371,840 bytes.
Release rose from 844,038,144 to 2,849,067,008 bytes. Native correctness and faster
timings cannot qualify this regression. The baseline source remains
`fac05d5b70b9a5933b12093c1b2337f76aee8cdc`; no threshold, sanitizer setting,
native watchdog or baseline is changed. No Full was started for this candidate.

`manifest.json` binds the complete failed comparison/state, six native verdicts,
complete native MFT metrics and sampled memory, both parser verdicts, the service
verdict, build receipts and exact external drivers. The installed MSVC 14.52
resource-header digest and the next proposed settings are recorded separately
in `pool-options-analysis.json` as generated analysis.

Inspection found that this implementation's default pool options admit very
large allocations and allow extremely large chunk growth. That is evidence for
testing explicit small-allocation/chunk limits, not proof of the next candidate's
performance. The next candidate must be compared with the same retained ordinary
allocator baselines and satisfy timing/memory admission before final qualification.
The remaining Close repair, six final builds, all Fresh Full profiles, ETW/ASan
canary and public migration remain pending. Failed receipts remain failed.
