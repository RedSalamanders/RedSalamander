# I26 ordinary-action baseline

Expected failing native witness, x64 Debug, 2026-09-19. This is focused evidence, not Full closeout.
The renderer is unchanged; only test observations, assertions and the fixed redraw workload were added.

- Run: `20260919T101114Z-99048-3bea933bc7d849c8a2694e3c87cf8c62`.
- HEAD: `a212edec433d658748ae1273cba68633c86de8b2`; dirty test/diagnostic inputs.
- Build receipt: `8c24cefa30b1e7038e9861a0d445473bd594b00e191feb443cd3a542bb45004f`.
- Source snapshot: `571772823c2bdc7c2ac27c5535506b31e380938e6d29cf1ed5e5a00583d7988e`.
- Toolchain identity: `ec868e067972f0bd3f7d27e3d565da1e7fc42e307eac23af2de0c6400f25e7db`.
- Executable SHA-256: `3EA533AE6D48CB4F40CE6501F89CB342C802068887D1B5E0D6139F3C7EE85EA5`.
- Phases05_06.cpp fixture SHA-256: `AC120236424259785F8BD265B46502C327F7DF04B47E715EBF1BF493870A3568`.
- Raw metric SHA-256: `C518FE3AE5A3658A7238D1657AEF682352E42992BD5038FB83B991B36AF2A3F7`.

Command from the repository root, with `RSBuildEnableTests=true`:

```powershell
./Tools/Run-AllTests.ps1 -Suite FileOps -CaseFilter Phase6_PopupSmokeResizeAndPause -Platform x64 -Configuration Debug -TestRoot C:/RedSalamander.Perf -TimeoutMultiplier 3
```

Build: zero warnings/errors and published attestation. Native Setup and Cleanup passed; the target
case failed with Pause=false/false, Cancel=false/false and Speed=false/false (positive area/hit target).
The trace, all declared case outcomes, aggregate audit and raw metrics are retained without filtering.

`FileOps.Popup.ActionsWitness.RedrawBatchUs` performs four warm-up redraws, then five batches of
16 synchronous redraws of the process-owned popup. Values in microseconds: **134593, 116594, 130145,
132722, 133484**; median **132722**, native window **144 DPI**. This is redraw-call duration, not
presentation FPS. No pointer movement or activation is used. Compare with the identical candidate
fixture, profile and machine; this baseline alone makes no performance claim.

The aggregate sandbox audit independently reports 228 issues: 225 unrelated pre-existing run
directories and three task-owned logs placed at the root. The logs were subsequently moved into
`C:/RedSalamander.Perf/evidence/i26-ui`; unrelated runs were preserved. The audit remains an explicit
environment gate and must not be represented as clean. The archive-format validator passed.
