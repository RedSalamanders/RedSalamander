# File Operations title and provider fallback — focused Debug receipt

Source base `fd82f895f` plus the bounded title/question/test/spec diff;
qualified DxUi pin `78b3de389a189c7f86f611787e0489fb6d474218`.
Test-enabled x64 Debug build receipt
`deb3e0b3a7efe78ea015df4f594fb2669ebd216dac31b3db70180d74d50015cb`,
zero warnings and errors. Commands cases
`cmd_pane_fileops_conflict_comparison_rows` and
`cmd_pane_fileops_issues_pane_primary_grid` each passed 1/1 with zero skipped.
Runner IDs: `20260921T170759Z-56232-8eb0272e21c64452a438039dfc8a75d6`
and `20260921T170900Z-20448-27d57a4022764e1db5badc715bb6ba58`.

[Three unchanged application captures and their hashes](../../../../Plans/WIP/FileOperationsUiReview_2026-09-16/implementation/2026-09-21-neutral-title-fallback/README.md)
cover French exact-path fallback and English/French live Issues titles.
Full runner logs, results and original screenshot artifacts were copied before
the next runner cleanup to `C:/RedSalamander.Perf/evidence/i26-ui/neutral-*20260921`.
The runner reported 260 disk-audit issues; no clean-disk claim is made.

The captured primary Issues grid still exposes part of the next row under the
selected tall row. Its separate DxUi acceptance/resource gate remains open.
The final integrated Full/Release, physical 96-DPI/AT and paired resource gates
are not established by these two focused tests.
