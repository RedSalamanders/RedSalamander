# I26 ordinary-action baseline — settled fixture v2

Expected failing witness on the unchanged production layout, x64 Debug, native 144 DPI. The fixture
waits for layout and selected-card height motion to settle, and checks the eligible primary action:
Skip while discovery is active, Pause afterward. It still requires positive bounds and correct hit targets.

- Runner: `20260919T102221Z-93744-e6b559bc54234934ad4b90d893d39ff8`; metric run `2026-09-19_102617`.
- HEAD: `a212edec433d658748ae1273cba68633c86de8b2`, dirty fixture/test observations and documents.
- Receipt: `57250eabdc28c1cae3892e4252586401008e8cf8f2301a4f01f651b70438d3c4`.
- Source snapshot: `d79a30fd548a7f1d5087492533163671db186138d0217d346f44fbddb766d703`.
- Fixture SHA-256: `91364B97CF190E44C930ED1C68E99C6832135EF243F8A1A346513DAB7A82AE58`.
- Executable SHA-256: `F18B4ED67610D0C8D4394EA7FA90EFA57A41700E97C99A32B121F8795058069F`.
- Raw metric SHA-256: `799650A3CAE65D636CEC807C8AE0F424F2282F122672B83A77C70AE1D1469B82`.

The native case fails with primary=false/false, cancel=false/false, speed=false/false, discovery=true.
Setup and Cleanup pass. Build: zero warnings/errors. Independent sandbox audit: 225 pre-existing run
directories, preserved. This focused expected failure is not a clean repository closeout.

Same command as the first baseline: `Run-AllTests.ps1 -Suite FileOps -CaseFilter
Phase6_PopupSmokeResizeAndPause -Platform x64 -Configuration Debug -TestRoot C:/RedSalamander.Perf
-TimeoutMultiplier 3`, with `RSBuildEnableTests=true`. Four warm-up redraws precede five 16-redraw
batches: **166907, 141477, 159969, 151136, 130296** microseconds; median **151136**.
These are synchronous owned-window redraw durations, not display FPS. Compare only with the same v2
fixture on this machine/profile/DPI; keep original attempts for diagnosis, not as interchangeable baselines.
