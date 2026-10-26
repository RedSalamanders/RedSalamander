# R1d common Preparing baseline

- Machine profile: `4cb089111a23`
- Source commit: `7f7d24888161bc36e67757be624bc86ff90bb414` (R1d RED contract)
- Build: test-enabled x64 Release; direct MSBuild completed with `0 warning(s), 0 error(s)`
- Command: `RedSalamander.exe --fileops-selftest --selftest-case=FileOpsFamily_R1dPreparingLifecycle --perf --selftest-no-activate`
- Runs: `r1d-preparing-baseline-clean-01` through `-05`
- Scenario: five sequential independent processes; each projects 4,096 selected top-level roots and 8,192 scope facts without descendant traversal or provider content I/O
- `FileOps.Preparing.BuildSnapshotUs`: `2,785`, `2,390`, `2,263`, `2,387`, and `2,405` us; nearest-rank p95 `2,785` us
- `FileOps.Preparing.SelectedRootReadinessUs`: `7,087`, `6,724`, `6,643`, `6,588`, and `7,433` us; nearest-rank p95 `7,433` us
- `FileOps.Preparing.UiAdmissionUs` process maxima: `1,025`, `1,071`, `893`, `908`, and `1,026` us; nearest-rank p95 `1,071` us
- Retained preparation projection: `2,736,128` bytes in every process

Every baseline process is intentionally RED (`2/1/0`): the legacy engine rejects the common pre-consumption gate for an ordinary Copy task. The retained baseline metrics isolate top-level preparation projection cost; they do not claim the missing lifecycle, legal phase order, or before-Ready mutation guarantees.
