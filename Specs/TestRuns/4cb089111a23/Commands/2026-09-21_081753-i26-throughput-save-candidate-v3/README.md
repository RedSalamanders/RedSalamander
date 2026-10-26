# Throughput save and reopen: throughput-save-candidate-v3

0 passed / 3 failed / 0 skipped; x64 Debug.
The fixture denies atomic replacement of the runner-isolated settings file, verifies
the live choice remains changed while disk retains its prior value, releases the
denial, then verifies explicit save recovery and both ordinary/expanded placement
keys. Absent settings correctly mean default-expanded. It recreates the popup with
the same running task and requires identical outer bounds and a collapsed summary.

V3 is the retained product regression: all three repetitions shrink the reopened
height by 72 pixels (632 -> 560 at native 144 DPI), with position and disclosure
correct. V4 preserves the restored initial expanded geometry; subsequent content
changes retain normal growth policy. The V4 prefix run additionally exercises the
existing popup progress, resource, failure-surface and presentation fixtures.
Consult actual results; no failed run is claimed to pass.

Earlier V1/V2 fixture failures remain externally retained: V1 incorrectly required
an explicit optional value for default=true; V2 compared different task contents.
Only V3 is the matched unchanged-product witness for the corrected reopen fixture.
Reopen is not process restart or external hot reload. Those, native mixed-DPI/AT,
Release and final Full qualification remain open.

Raw: `S:/RedSalamander.Perf/evidence/i26-throughput-save-candidate-v3-20260921`. Exact source/build snapshots agree. The compact throughput
redraw metrics retain original line provenance; prefix-run timing has broader test
order than the baseline, so report variation rather than overall non-regression.
`captures.json` records raw harness PNG hashes without republishing unreviewed images.
