# File Operations graph transparency Release candidate

This run validates the shared DxUi throughput-graph painter after restoring translucent
per-stream band fills. It is candidate-only evidence: the older accepted graph run uses a
different machine profile, so this archive does not claim a before/after speedup.

## Source and command

- Git base: `fa2ca6a4e8e6`
- Tested worktree patch identity (`git diff --binary | git hash-object --stdin`):
  `c902b6738b7904b5b09fbced42373d83c611bb72`
- Configuration: x64 Release with `RSBuildEnableTests=true`
- Command: `REDSALAMANDER_DXUI_GRAPH_PERF=1 .build/x64/Release/DxUiTests.exe --suite=Rendering --perf-jsonl=<run>/perf_metrics.jsonl`
- Result: `Rendering` passed, exit code 0.

The repository build wrapper compiled `DxUiTests.exe` with 0 warnings and 0 errors. Its
receipt publication remained unavailable because the wrapper classified the manifest-mode
vcpkg `blake3.dll` source path as outside the selected output directory.

## Bounded-work result

Five deterministic 180-sample by 16-slot captures reported:

- capture median 33,661 us; mean 32,871.8 us;
- band-painter median 111 us; mean 115.2 us;
- 2,864 logical quads and one rendered geometry in every capture;
- 16 active color slots in every capture.

`perf_metrics.jsonl` retains the full 1,325-row runner output so the claim-bearing rows remain
in their original order. SHA-256:
`26DB87B9D8BB47A1947D179B357D295FB8E8B1B0CA97CA4023FB9852D5250E1D`.
