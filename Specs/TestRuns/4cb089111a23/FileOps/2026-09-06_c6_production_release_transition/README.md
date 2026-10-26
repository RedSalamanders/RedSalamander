# C6: production Release transition

The real test-enabled-to-production Release transition passes: 13 test-only
binaries retire, 164 production receipt artifacts verify, both shipping SQLite
copies remain, and the retained package-contract host runs successfully.
This closes that C6 prerequisite, not the core plan or final Fresh Full gate.

## Exact source and build identity

Build/test source was clean committed
`0ffe4e18419ad10a8db43c58afa626760b5e0852`. The source/index stayed frozen through
the build and both contained smoke runs. Only plan/archive/attribute changes
followed testing. No implementation repair was necessary for this transition.

Before transition, the canonical receipt loader verified the test-enabled pointer
and immutable record, and all 177 artifact sizes/digests were checked under the
Release artifact-profile lease. That prior runtime baseline is the earlier Clean
checkpoint, not a claim that the old receipt matches the later diagnostic commit.

| Identity | Before | After |
|---|---|---|
| Tests enabled | true | false |
| Receipt | `a6fece6675db82bf539822d2caf671a416ad4918b27e0c406522c637dff2576d` | `6af048447fb32e79182634085a2ee3367f0234f31c3225567e2588c57eef6341` |
| Source snapshot | `3754f54aea44516549094354eff5053a2588e68eb2fc6d939f6fec34518b7335` | `dd4b8a8b40736ba27199adad66d858f7d91b31a6507fb9233522a511eb255652` |
| Verified artifacts | 177 | 164 |

`build.ps1 -Configuration Release` received explicit `RSBuildEnableTests=false`
and forced the required full rebuild because the old receipt was incompatible.
The original process override was restored in `finally`. Profile/dependency
coordination, parallel MSBuild, required runtime staging and fail-closed receipt
publication remained enabled. No warnings or errors were reported.

## Transition and executable checks

`verification.json` retains all 21 named checks and their observed counts/results.
All pass. Each retired binary existed in the before receipt, is absent from the
after receipt, and is absent on disk:

- CrashHandlingTests, DxUiTests, FileSystemCurlTests, LocalizationTests, MonitorTest,
  RedConfigureTests, SettingsSchemaTests, ViewerPETests and ViewerSqliteTests executables;
- PerformanceTests2.dll;
- the cs-CZ, fr-FR and sk-SK LocalizationTests resource DLLs.

These are generated test outputs, recoverable by a test-enabled rebuild. No user
files or old test evidence were removed. The production application, Search Service,
root/Plugins SQLite DLLs and PluginContractTests.exe remain attested. The Search
Service's runtime closure explicitly contains root `sqlite3.dll`.

The contained Search Service `--help` run exits zero; it does not register, start
or compact a service/store. The contained `PluginContractTests.exe --package-smoke`
run exits zero with 1,365 OK assertions, zero failed/skipped assertions and the
terminal package-smoke pass line. This is an in-place profile smoke, not clean
ZIP extraction, a live-provider mutation test, or a final Full result.

All 164 after-receipt artifacts were independently size/digest checked immediately
before smoke under the Release profile lease. The canonical loader authenticated
the pointer against its immutable receipt record. No ordinary connection settings,
Windows Hello policy, TLS/SSH policy or live remote data were changed.

## Raw evidence and timing

`copy-manifest.json` binds the seven byte-for-byte copies of the before/after
receipts, verification records, build log and two smoke logs. Raw baseline data
remain under `C:\RedSalamander.Perf\runs\c6-production-transition-20260906`; smoke
data remain under `c6-production-transition-verify-20260906` at the same root.
The preserved command texts record exactly what ran. Reproduction must use new
governed run IDs and a verified test-enabled baseline.

The command's final console projection mistakenly applied `Select-Object` directly
to an ordered dictionary and displayed null summary fields. The saved JSON was
written correctly before that projection and was independently reopened to verify
21 passed/0 failed, the exact receipt and all named outcomes above. No test was
rerun or failure reclassified to correct that presentation.

`perf_metrics.jsonl` derives three rows from the exact receipt/verification records:
164,299 ms build duration, 142.4598 ms Search Service help, and 437.3845 ms package
smoke. These are single-run wall-clock observations, not a before/after throughput
claim or a universal timing limit. They include process-launch overhead for smoke.

## Continuation

Validation: all seven raw-copy identities, 21 outcomes and three derived timing
rows were independently checked. Build-log replay finds zero warnings/errors.
Archive checks pass for 12 pre-stage files, 2,208 changed-set files and 3,421
whole-inventory files; all 12 staged archive files preserve their raw byte identity.
The specification inventory reports zero blocking findings.

Return to the core plan's remaining C0 provider/memory/route requirements, followed
by C1–C5 and final consolidation/Fresh Full/inherited I17. Release currently has
test-disabled outputs. A new compatible test-enabled full-solution receipt is
required before native FileOps selftests. Do not treat production smoke as their
replacement, change the four deferred features back into core, or move the plan
to Done yet.
