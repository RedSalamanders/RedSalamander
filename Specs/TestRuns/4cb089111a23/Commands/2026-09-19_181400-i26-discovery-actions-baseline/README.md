# French discovery action — retained baseline

Run `discovery-actions-retained-baseline` fails 0/1, zero skips, because the French label needs
184.48828 pixels but the control receives 126 pixels at native 144 DPI. The geometry witness covers
480/640/760/480-DIP widths, active action targets and overlap independently of production allocation.
Build: zero warnings/errors. Sandbox audit: zero issues.

Receipt `300ea53a85f31c9aa9fab196856762c457103c271ab95b30c6230031b353a249` binds workspace
`ae337f0713bce9a2bea249c6a58c906cb301d3619467bee491b068babcc56a21`. Fixture SHA-256 is
`e60006635e0c358fae90c729dfc5a0df863ed574dcccd47e655fb2fc27707c2e` in both runs. This retained
baseline precedes restoring the production fix. An earlier automatic run was removed by sandbox
cleanup and is not used as paired evidence; the first candidate had one test conversion warning,
corrected before both retained runs.

All 20 `FileOps.Popup.DiscoveryActions.RedrawBatchUs` rows are retained, each eight redraws;
five batches per width. Median 66,430 us, nearest-rank p95 67,065 us. These are synchronous Debug
diagnostics, not frame pacing or Release acceptance. Compare the [candidate](../2026-09-19_181400-i26-discovery-actions-candidate/README.md).

Raw run files are preserved in `S:/RedSalamander.Perf/evidence/discovery-actions-retained-baseline`;
logs and a second artifact copy are in `C:/RedSalamander.Perf/evidence/i26-ui/discovery-actions-retained-baseline*`.
Only trailing runner-log whitespace was normalized in the archive.
