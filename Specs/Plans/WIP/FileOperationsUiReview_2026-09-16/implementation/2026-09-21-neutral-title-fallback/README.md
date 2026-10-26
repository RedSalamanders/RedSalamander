# Neutral Issues title and unknown-provider conflict question

These are unmodified **application test-harness captures** from x64 Debug, source
base `fd82f895f` plus the reviewed working-tree changes. The build was attested
as `deb3e0b3a7efe78ea015df4f594fb2669ebd216dac31b3db70180d74d50015cb`
with zero warnings/errors; DxUi remains pinned to
`78b3de389a189c7f86f611787e0489fb6d474218`. The
[manifest](capture-manifest.json) records the exact source, run and SHA-256 of
each PNG. [Open the focused gallery](gallery.html).

The unknown-provider collision now states only `La destination existe déjà.`;
the exact `/photographies/été\original.raw` remains in the labelled destination
row, with no invented leaf or parent. The Issues native title now describes
issues rather than asserting that every row failed. The same live-language
fixture includes pending prompts, an accepted `Remplacer` decision, legacy
diagnostics and a later completed result. Its English and French title refresh
and the fallback UIA assertions pass in the focused Commands harness.

| Focused harness case | Result | Run ID |
|---|---|---|
| `cmd_pane_fileops_conflict_comparison_rows` | 1/1 passed | `20260921T170759Z-56232-8eb0272e21c64452a438039dfc8a75d6` |
| `cmd_pane_fileops_issues_pane_primary_grid` | 1/1 passed | `20260921T170900Z-20448-27d57a4022764e1db5badc715bb6ba58` |

Full runner logs and copies of native artifacts live under
`C:/RedSalamander.Perf/evidence/i26-ui/neutral-*20260921`. The fixture images
were visually inspected. The Issues primary grid still shows a sliver of the
next tall row at its lower edge; its separate DxUi row-layout/resource gate is
open. These focused runs do not establish final Full/Release, native 96-DPI,
assistive-technology or paired resource qualification. Historical captures and
their labels remain unchanged.
