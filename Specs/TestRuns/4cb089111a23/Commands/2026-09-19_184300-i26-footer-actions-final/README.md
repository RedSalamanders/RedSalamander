# French footer — final candidate

Run `footer-actions-final` passes 1/1, zero skips, with zero warnings/errors and zero sandbox issues.
All required running-copy footer controls remain reachable and non-overlapping at 480/640/760/480
DIP, with at least 24 DIP of label padding. Three final French harness images and one baseline image
were [visually reviewed](../../../../Plans/WIP/FileOperationsUiReview_2026-09-16/implementation/2026-09-19-footer-action-fit/README.md).

Receipt `9e2a6253d91cb47f5f666b61816b3cf686b62388e1b6632d87aeae345ba0af1e` binds workspace
`e52e828dc55849b449f521445a514c68775c116f8d6799c8ae45fda76b7f046e`. Both full metric streams
and the unchanged test source hash are retained; [paired-summary.json](paired-summary.json) preserves
all ordered redraw samples and identities. The first candidate passed the test but had four repeated
C5246 aggregate-initialization warnings. Corrected braces precede this final zero-warning build.

| Eight-redraw diagnostic | Baseline | Final |
|---|---:|---:|
| Count | 20 | 20 |
| Median, us | 66,405.5 | 66,561 |
| Nearest-rank p95, us | 67,320 | 67,160 |

Median changes +0.234%; p95 -0.238%. The candidate records two three-label cache fills (107/80 us)
before the measured sequence, with no repeated text measurement during its 160 redraws. One
three-string vocabulary, two logical widths and one format identity are retained per popup. This
pair shows no material synchronous Debug redraw change; it does not establish presented FPS,
allocations/idle budgets, Release, native mixed DPI or full paused/Resume/keyboard/UIA coverage.

Raw logs/copies remain at `C:/RedSalamander.Perf/evidence/i26-ui/footer-actions-final*`; the completed
run was preserved under `S:/RedSalamander.Perf/evidence/footer-actions-final`. Archived runner-log
whitespace alone is normalized. General wrapping and footer height remain active implementation work.
