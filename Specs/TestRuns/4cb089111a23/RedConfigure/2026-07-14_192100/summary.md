# RedConfigure manager closeout evidence

- Date: 2026-07-14
- Profile: `4cb089111a23`
- Build: Debug x64, 0 warnings / 0 errors
- Command: `build.ps1 -ProjectName RedConfigureTests`, then `RedConfigureTests.exe` with `REDCONFIGURE_PERF_JSONL`
- Result: passed all 39 standalone cases
- Windows SDK resource compiler: available; generated `.rc` accepted and `.res` produced

## Deterministic repo-sized scenario

- 6 resource owners
- 1,500 English source rows
- 4 target cultures (6,000 target cells)
- 10 theme files
- Complete scan/theme load/localization parse: 320 ms (target < 500 ms)
- Combined validation: 3 ms (target < 250 ms)
- Theme edit plus preview-model resolve: 344 µs (target < 16 ms)

`perf_metrics.jsonl` contains the complete focused run, including workspace discovery, theme catalog load, per-file localization parse, review load, validation, and theme-preview resolver metrics. The same executable run also covered four-page workbench creation/switching, generated-RC compiler acceptance, batch/clipboard/undo workflows, grouped theme export, and atomic post-write reparse.
