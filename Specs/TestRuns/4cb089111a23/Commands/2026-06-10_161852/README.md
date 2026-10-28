# Batch Rename Preview Perf Candidate

## Command

```powershell
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_batchRename_
```

## Environment

- Date: 2026-06-10
- Machine profile: `4cb089111a23`
- Configuration: Debug x64
- Branch: `batch-rename-window`
- Commit at run start: `c468032f0cfb9a29ab3018adbec231548f6e57c7`
- Result: 17 passed, 0 failed, 0 skipped

## Scenario

`cmd_pane_batchRename_engine_large_preview_perf` builds a deterministic 10,000-target preview using counter macros, date/time macros, literal search/replace, extension exclusion, and Change Case. The case asserts the preview metric family is emitted and checks the emitted `batchrename.preview.build_plan_us` duration against an explicit 5,000,000 us upper bound.

## Metric Summary

- `batchrename.preview.build_plan_us`: 447,396 us for the 10,000-row macro scenario (`value0=10000`, `value1=10000`).
- `batchrename.preview.rows`: 10,000
- `batchrename.preview.changed`: 10,000
- `batchrename.preview.errors`: 0
- `batchrename.preview.warnings`: 0

## Threshold

The current guard requires deterministic correctness, emitted perf metrics, and `batchrename.preview.build_plan_us <= 5,000,000` for the synthetic 10,000-target macro scenario. Future closeout work should add same-machine baseline/candidate comparisons for the remaining preview, manual, collection, and execution scenarios.
