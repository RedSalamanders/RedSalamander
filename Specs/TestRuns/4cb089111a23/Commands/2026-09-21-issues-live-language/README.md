# Existing Issues pane: coherent language changes

Machine `4cb089111a23`, x64 Debug, base commit
`41cca99b9a54c84c4cbf94e2be7545a16ad015db` plus each runner's bound dirty source
snapshot. Qualified DxUi pin `78b3de389a189c7f86f611787e0489fb6d474218` is unchanged.
All stages use the same `cmd_pane_fileops_issues_pane_primary_grid` fixture with
three repetitions. Each repetition changes the same window en-US then fr-FR in
primary and technical presentations, followed by five unchanged theme notifications.

| Stage | Passed / failed / skipped | Language notification median | Five unchanged notifications median |
|---|---|---|---|
| baseline | 0 / 3 / 0 | 4.450 ms | 2.220 ms |
| candidate | 3 / 0 / 0 | 9.617 ms | 25.690 ms |
| optimized | 3 / 0 / 0 | 9.277 ms | 7.135 ms |
| revision | 3 / 0 / 0 | 8.982 ms | 2.011 ms |

The baseline mixes old French captions/headings with new English detail labels.
The first candidate repairs title, column titles, row labels, detail, actions,
copy feedback, empty-state text and accessible detail name, but repeats unnecessary
work for unchanged theme messages. The intermediate optimization avoids replacing
unchanged titles and republishing unchanged grid accessibility. The final revision
uses the localization manager's synchronized successful-application revision to skip
translation/row reconstruction on unrelated theme messages.

Final language transitions publish exactly one changed row projection per
notification; all 60 unchanged notifications publish none. Stable selected row,
logical grid focus, semantic sort, disclosure, visible column identities, positions
and widths survive every tested transition. The existing saved-layout/reopen checks
also pass. The original French diagnostic sentence and provider paths remain their
recorded facts; localizing a presentation does not translate historical messages.

There are 12 timing samples per stage/metric. `metrics-summary.json` includes all
samples' source hashes, medians and nearest-rank p95/p99. The final real language
change performs more work than the defective baseline (about 4.53 ms additional
median), while unchanged-notification median returns to baseline range. These are
notification costs, not presented-frame latency. The final unchanged-batch maximum
is 48.498 ms versus 24.215 ms in the baseline; these small, variable samples do not
establish tail-latency non-regression or the broad I26 performance/resource gate.
No memory or GPU non-regression is inferred from timing or functional passes.

The new Common API has focused native localization tests for successful application,
same-language reapplication, rejected preference and retained active resources;
`localization-tests.log` preserves the complete successful suite output. The final
File Operations Commands family passes **50/50**, no skips. Its aggregate remains
in `family`; raw captures/logs are retained under
`C:/RedSalamander.Perf/evidence/i26-ui/issues-live-locale-family-20260921`.

All four full test-enabled x64 Debug builds have zero warnings/errors.
`build-receipts.json` binds each build log and receipt. The final revision receipt is
`0135cb7402b71af67df1c3d713bd3b959d128412ef6a8127847d1bc0ad52e7b9`.
Reviewed actual application images and capture provenance are in the linked I26
implementation gallery. Captures use the deterministic application harness; no
desktop screenshot or proposal rendering is used as application evidence.

Runner disk-audit findings are retained (255 baseline through 259 family), without
a clean-sandbox claim. This archive does not close Full/Release, complete native
mixed-DPI/AT, large-model resource or other I26 gates. The separate described-menu
memory acceptance is unchanged and no library branch has been adopted implicitly.

**Full Debug qualification failed:** generated run
`20260921T144647Z-81208-b80aa09981ff4858af41b888e8ff80e5` exited 1.
The aggregate records 1997 passed, 2 failed and 37 skipped cases, with 19 promoted
entries and one incomplete entry (FileOperations). Commands passes 912/0/2,
including all 50 File Operations UI cases; both Commands skips are declared opt-ins.
The tooling-driven full rebuild has zero warnings/errors, receipt
`cdd541db0ebbb420296a0c4a3d34d048e26b089c3e18f39dc3eb5c2131fb4eb3`.
FileOperations failed `Floodgate_InlineF2WorkerQueueBypassAndInterlock` (180046 ms)
and `Phase5_DiscoveryCancelReleasesSlot` (703 ms), then stopped advancing at
`R4A19_DiscoveryProviderControls` for over six minutes. The owned no-activation
test process 90424 was terminated only after retaining its identity, traces,
metrics and a live minidump; this is interrupted coverage, not a successful suite
or a proven crash. The independent Release application was untouched.
Raw evidence remains in
`C:/RedSalamander.Perf/evidence/i26-ui/issues-live-locale-full-20260921`;
the complete runner log is
`C:/RedSalamander.Perf/evidence/i26-ui/issues-live-locale-full-generated-20260921.log`.
The failures remain open; focused localization passes do not waive Full.
