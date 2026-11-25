# I66 clean Completed group

The x64 Debug no-activation `FileOps_ClearCompletedScope` run
`20261124T202452Z-13956-4624655892f14dfe948ad1ea1465752b` passed setup, the exact case and cleanup.
The selected app build emitted zero warnings and zero errors. The verified receipt was
`b4604108de105e5db7dd87b3956c94a28ea1883061f2682deacf71c8e8b27572`.
Both collapsed and expanded Clear completed removed the two clean task IDs shown at the last paint.
It retained failed, partial, uncertain and warning results, an interrupted-Move recovery notice and
the clean result added after that paint.

The same-machine x64 Debug `FileOps_PopupIdleCadence` pair passed setup, the case and cleanup
on both variants without desktop activation. The baseline uses the merged corroborated Rename tree at
`8791d9756` with a verified build from its byte-identical PR head. The candidate adds clean grouping.
These are single observations within the existing cadence budgets.

| Variant | Run | Synthetic idle ticks / snapshots | Sole running engine ticks / snapshots | Paused engine ticks / snapshots |
|---|---|---:|---:|---:|
| Baseline | `20261124T202656Z-100300-5e08bc1c39b74a02b0d1312b4598cbd1` | 0 / 0 | 11 / 27 | 2 / 2 |
| Candidate | `20261124T202844Z-100300-e1cdea8121df49719c027d7df1857700` | 0 / 0 | 11 / 28 | 2 / 2 |

The retained files contain the ten `FileOps.PopupCadence.Counters` records from each source trace
in original order. The complete source files remain in the governed sandbox archives.

| Variant | Original metric rows | Original bytes | Original SHA-256 |
|---|---:|---:|---|
| Baseline | 7303 | 2654992 | `996e73c339dc4e0b7f073917d15bcf98a77a79ed409bedf0409df1e630ab79c4` |
| Candidate | 7218 | 2609636 | `99ff7f5589bec86b63393dc6f16e7ac3d1c2cbf7368bfe0dcd43c59a998d994a` |
