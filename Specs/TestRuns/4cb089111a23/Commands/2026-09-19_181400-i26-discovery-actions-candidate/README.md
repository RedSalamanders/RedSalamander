# French discovery action — candidate

Run `discovery-actions-retained-candidate` passes 1/1, zero skips, with zero build warnings/errors.
The full French label plus 24-DIP padding fits at 480/640/760/480 DIP; Skip, limit and Cancel retain
active hit targets without button overlap. Three [harness images](../../../../Plans/WIP/FileOperationsUiReview_2026-09-16/implementation/2026-09-19-discovery-action-fit/README.md)
were visually reviewed. Narrow footer/path defects remain outside this bounded repair.

Receipt `cf1fdd8325ca006ce167d60b78540e819798efd560f1aa02173ac023898e7ad6` binds workspace
`e27ee2e16e041028a512209b4f81ae2436a5342334a253b2dae2a9bad5c8f30f`. The unchanged fixture hash,
ordered values and both identities are in [paired-summary.json](paired-summary.json); full metrics
are retained. Same machine, dark theme, native 144 DPI, x64 Debug and synthetic running-copy data.

| Eight-redraw diagnostic | Baseline | Candidate |
|---|---:|---:|
| Count | 20 | 20 |
| Median, us | 66,430 | 66,422.5 |
| Nearest-rank p95, us | 67,065 | 67,135 |
| Production label measurements | Not implemented | 1 (14 us) |

Median changes -0.011%; p95 +0.104%. No material redraw change is observed in this pair; no broad
performance conclusion is claimed. The cache retains one label and one format per popup. Release,
allocation/idle budgets, presented FPS, keyboard/UIA and fresh physical mixed DPI remain separate
gates. The independently launched Release app was preserved.

The runner's original audit reports one issue for this task's preserved named baseline directory.
Both completed runs were then moved under the sandbox's evidence folder; the separate
[post-archive audit](post-archive-audit.json) reports zero issues. This does not rewrite the runner
verdict or imply another test run. Raw run files remain in
`S:/RedSalamander.Perf/evidence/discovery-actions-retained-candidate`, with logs/artifact copies at
`C:/RedSalamander.Perf/evidence/i26-ui/discovery-actions-retained-candidate*`. Runner-log trailing
whitespace alone was normalized.
