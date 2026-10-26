# Speed dialog — final repeated candidate and comparison

`speed-layout-candidate-repeat` passes 3/3, zero skips, with zero build warnings/errors and a clean
sandbox audit. Live French label/input/help/actions fit in all four states; the existing UIA surface,
ValuePattern, confirmation and nested-shutdown checks also pass. Five final/baseline
[harness images](../../../../Plans/WIP/FileOperationsUiReview_2026-09-16/implementation/2026-09-19-speed-dialog-layout/README.md)
were reviewed. Final source adds a minimized/empty-client guard after the first candidate.

Receipt `4c4d04b8bfdf3ecc948c8bb1cddab676faf040347e1b8f6c86c17c893ed340a2` binds workspace
`20e929da89cbba329de8af7f88ed73376ecdd25b57e07fb4323987c07fc5a68f`. Every variant uses test SHA-256
`aa40029bddbb4454a9caf8a421bce868e739f2fe032d22fae476d5cf230f7f26`. [All series](all-series.json)
retains identities and every ordered sample; each sibling archive preserves complete raw metrics.

| Variant | Samples | Median, us | Nearest-rank p95, us |
|---|---:|---:|---:|
| Original baseline | 20 | 66,527 | 102,916 |
| First candidate | 20 | 66,657.5 | 118,309 |
| Repeated baseline | 60 | 66,509.5 | 121,362 |
| Final repeated candidate | 60 | 66,514.5 | 112,266 |

The first pair's p95 increase is retained. It did not repeat: the repeated medians differ by 5 us,
and candidate p95 is about 7.5% lower. The unchanged baseline's tail also varied. These runs do not
establish a consistent degradation, and the latter pair does not erase the first flag or prove
universal non-regression. Each repetition measures initial, constrained, validation and corrected
states in that order, five batches of eight redraws each. Candidate geometry intentionally grows to
fit the same constrained request; rendered pixel areas differ. Whole-case durations are not paired:
the expected baseline failure bypasses the later shutdown check. No samples are filtered or pooled.

This is x64 Debug, French, native 144 DPI, the Acrylic tool-window selftest theme, and the governed
foreground modal harness. It does not qualify presented FPS, Release, idle/allocation budgets,
physical mixed DPI, extremely short work areas, live locale switching or all keyboard interactions.
All broader gates remain open.

Raw artifacts remain under `C:/RedSalamander.Perf/evidence/i26-ui/speed-layout-*` and the corresponding
`S:/RedSalamander.Perf/evidence` run directories. Hash-verified baseline source and final source are
saved in the C: evidence folder's `speed-layout-source/baseline` and `speed-layout-source/final`.
Only trailing runner whitespace is normalized in this archive.
