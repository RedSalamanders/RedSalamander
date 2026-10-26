# File Operations card hierarchy closeout

2026-10-16, machine `4cb089111a23`, native x64 Debug, base commit
`761c6b622972d49f1674ff2fd322ea0aab738e0b` plus `source-diff.txt`.
Both final builds have source snapshot
`c4566839c3d9746d8e474d9fffca58f17f58dc50a04a547e8e30944722380153`.
These receipts bind the implementation before the pre-merge pinned formatter
normalized the touched native files. That normalization changes spacing, line
breaks and adjacent string-literal wrapping; merge checks have their own commit
identity. The final documentation screenshot/provenance and completed-plan
closeout were also added after these captured runs.
This is scoped qualification. Full, Release, ARM64, ASan, foreground-input,
other DPI, real SMB performance and long-run memory retention were not evaluated.

## Passing runs

| Case | Governed run | Build receipt |
| --- | --- | --- |
| `FileOps_ReviewRemediationContracts` | `20261016T173738Z-83008-21037ccaa05048f38559cb8314746aa0` | `c6667ebbe6f3397737f731ef836a606ef6dbb0726b6e00838e2c7532b50d19b1` |
| `DiscoveryScope_FolderLiveProgress` | `20261016T174403Z-41384-034b5e0cebd6459b8554f6aa0b1263d5` | `41962790722c6f3049d8d5d204dcaa22dc32d21545c5eb10358cca253426e40b` |

Each command was `Test-Changes.ps1 -Suite FileOps -CaseFilter <case> -TestRoot C:\RedSalamander.Perf`.
Each passed its three reported phases (setup, named case, cleanup), with zero
failures/skips. The application builds completed without compiler warnings/errors.
Each run's native results, runner report, trace, build receipt and build log are
retained in `layout/` and `live/`. No separate-build receipts are combined into
a Full qualification. `contracts.xml` records 214/214 passing localization and
harness-source checks; the runner also accepted all 118 native test filenames.

The runner reports remain historically unclean for the test disk audit: five
unrelated existing paths, plus the completed concurrent Pester scratch directory
in the live run. That owned Pester directory was subsequently removed with
`Remove-RSTestSandboxExactRunDirectory`; unrelated paths were left untouched.
No repository-wide green or clean-disk result is claimed.

## Native harness screenshots

All 80 PNGs are unchanged captures of owned native application windows, with no
activation or desktop-global input. The deterministic layout fixture renders
production cards from snapshots; it never admits operations on its displayed
`C:\Movies` / `D:\Movies` paths. These are actual application renders, not mockups.

`captures/layout/card-hierarchy/` contains 60 captures. `scenarios.tsv` identifies
the 52 base scenes: 13 states, English/French, 480/960 DIP, actual 144 DPI,
light at 480 DIP and dark at 960 DIP. The additional four `technical-*` and four
`throughput-collapsed-*` frames identify their language/width in the filename.

- [Expanded Move](captures/layout/card-hierarchy/scene0-en-US-960dip.png)
- [Narrow French card](captures/layout/card-hierarchy/scene0-fr-FR-480dip.png)
- [Collapsed failed Move retains its filename](captures/layout/card-hierarchy/scene3-en-US-960dip.png)
- [Collapsed mixed selection](captures/layout/card-hierarchy/scene5-en-US-480dip.png)
- [Technical details open](captures/layout/card-hierarchy/technical-en-US-960dip.png)
- [Plot hidden while rate/ETA remain visible](captures/layout/card-hierarchy/throughput-collapsed-fr-FR-480dip.png)

Assertions cover original root identity across 11 statuses and changing child
totals; actual painted subject/extension; paired fact geometry; native technical
caption fit and expanded state; non-overlapping action rows; disclosure choices
across card collapse; graph-label bounds; unknown-byte Copy without a percentage;
known zero-byte copied-file progress; item-based Delete; conflicts; Unicode; and
the ten queued missing-source file/folder fixtures.

`captures/live/folder-copy-progress/` contains 20 real-operation captures. The live
case passes 64 assertions across local serial, overwrite and parallel copies,
simulated-provider parallel and verified bridges, and open discovery. It checks
exact file/folder/byte totals, visible counter ink in Technical details, actual
per-file track-to-fill pixels, terminal lifecycle facts, and independent content
of every destination file (including empty folders). Counter captures reveal the
actual owned scroll rectangle, then restore the transfer view; no focus is taken.

## Rendering characterization

`render-samples.csv` retains the unrounded microsecond observations for the six
metrics in `render-summary.json`, from the layout run above. Percentiles use
nearest rank. The frame batch is 12 warmed redraws per base scene (624 samples).
The archive passed `Test-TestRunArchive.ps1 -RunPath <this directory>` with the
standard 2 MiB per-file and 5 MiB per-run limits; no limit overrides were used.

| Metric | Samples | Median | p95 | Maximum |
| --- | ---: | ---: | ---: | ---: |
| Warmed harness frame | 624 | 8.099 ms | 8.497 ms | 10.855 ms |
| Snapshot build | 724 | 0.029 ms | 0.045 ms | 0.087 ms |
| Card layout | 724 | 0.133 ms | 0.288 ms | 1.662 ms |
| Draw | 724 | 0.583 ms | 1.768 ms | 4.848 ms |
| Synchronous render total | 724 | 1.929 ms | 6.069 ms | 16.761 ms |

There are 56 fact-pair measurements across 724 renders; unchanged fact layouts
are cached. These observations characterize synchronous UI work and capture
preparation, not storage throughput. The original baseline's screenshots and
results were retained, but its perf directory was omitted before governed raw-run
cleanup. There is no matched pre-layout performance comparison, and differing
candidate fixture populations must not be compared as a causal improvement.

## Retained earlier attempts

The local ignored archive at
`../2026-10-16_card-hierarchy/` retains full raw metrics, compact harness archives,
source diffs and all failed candidates. It is deliberately separate from this
bounded closeout archive. It includes the original baseline; a corrected test
member/localization build failure; the caption-less chevron disclosure failure;
the incorrect unknown-size percentage assertion; and two live pixel-witness
failures caused by scrolling to the list bottom rather than the active counter.
Their original failed verdicts are retained. Later passing runs establish only
their own stated predicates and identities.
