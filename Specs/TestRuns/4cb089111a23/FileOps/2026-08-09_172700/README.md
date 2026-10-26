# Cinderstar Package 19 — bridge MOVE child-conflict cleanup

- Date: 2026-08-09
- Machine: `4cb089111a23`
- Build: x64 Debug, branch `operation-cinderstar-review`
- Build log: `.build/logs/msbuild-20260809_172237_635.log` (0 warnings, 0 errors)
- Expected disposition: `HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY)` (`0x8007012B`)

## Scenarios

| Scenario | Exact case | Result | Duration | Configured concurrency | Conflict prompts | Manifest peak / remaining |
|---|---|---:|---:|---:|---:|---:|
| Serial | `Fairstream_BridgePerFileConflictSkips` | Pass | 313 ms | 1 | 1 | 2 / 0 |
| Configured-parallel | `Cinderstar_BridgeMovePerFileConflictSkipsParallel` | Pass | 359 ms | 4 | 1 | 2 / 0 |

Both cases perform a real dummy-to-local cross-filesystem MOVE with a nested destination collision. They assert exactly one Skip prompt, the partial-copy result, unchanged skipped source and destination bytes, deletion of the successfully moved sibling source, and exact sibling destination bytes. The native policy assertions also require the full-success cleanup gate to retain `skippedFileConflictCount == 0` and allow partial manifest cleanup only for file-conflict skips—not reparse skips or unrelated failures.

The `serial/` and `parallel/` directories retain the unified-runner results, focused FileOps result and trace, and complete JSONL performance stream.
