# Deferred ASan qualification at d404fc90

The user deferred ARM64 and ASan on 2026-09-13 and will check them later.
These original local results remain evidence; this is not a successful Fresh Full run.
RedSalamander source is `d404fc90bcaf7793da4289ffcbd23debd6ee73b2`, DxUi pin `13788e95`.

- The x64 ASan Debug root build passes in 4m21s with zero warnings/errors.
- The seeded heap overflow exits nonzero with the expected sanitizer diagnostic.
- Pending-content and removal-focus each pass ten repeats. All 14 module identities
  are verified after each run. This validates the focused lifetime correction against
  the retained earlier heap-use-after-free reproduction.
- Bandwidth has 29 passes and one failure, including setup/cleanup. Repeat 10 reports
  cancellation latency 765,000 us versus 250,000 us. The driver stops; no later focused
  case or Full run is claimed. The previous duration failure is not reported here.

`case-comparison.json` preserves exact run IDs and the failing native case record.
`retained-files.json` binds original receipt bytes and locations. Complete archives
remain beneath `Z:/RedSalamander.Perf/evidence/I19-20260909`.
Resume through the [deferred qualification owner](../../../../Plans/WIP/DxUi_DeferredPlatformQualification_2026-09-13.md).
