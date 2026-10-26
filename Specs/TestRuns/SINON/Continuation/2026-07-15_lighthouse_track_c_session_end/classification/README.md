# Lighthouse Track C — Full failure classification

The Track C Full run exposed five failures outside the session-end implementation. Focused classification on
the same Debug x64 build produced the following results.

| Full failure | Focused classification | Result |
|--------------|------------------------|--------|
| Preferences reverse keyboard navigation | Exact Commands case, run `20260715T092433Z-65140-f5dd821388d24fa9828440da758eea18` | PASS, 1/1 |
| Tools payload source contract | Complete `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` | PASS, 118/118; exact assertion 91 ms |
| LocalizationTests access violation | Independently reproduced during Track B | External blocker; unchanged |
| RedConfigureTests access violation | Independently reproduced during Track B | External blocker; unchanged |
| Monitor ETW latency | Existing Full-suite baseline | External blocker; unchanged |

`commands_reverse_navigation/` contains the authoritative runner manifest and result for the exact Commands
rerun. Track B's durable blocker record is in
`Specs/TestRuns/SINON/Continuation/2026-07-15_lighthouse_track_b_full/`.
