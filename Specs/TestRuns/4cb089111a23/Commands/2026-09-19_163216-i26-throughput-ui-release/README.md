# Throughput UI — test-enabled Release

Run `20260919T143216Z-46632-1a1142be55de4771a63f18dafe6c7349` passes 1/1, zero skips,
zero disk-audit issues. Build: zero warnings/errors. Receipt
`34621d0175fee7287a3b83ea536f1f322a7a5a70d3c19a808b2e1eca68ad8369`, source snapshot
`222eaa722b78a7d4d4fe094ed75a92b01d7aa8f809d097482fcf0fd93921da1b`.
Set `RSBuildEnableTests=true` explicitly for the focused Release Commands runner.

The state/persistence assertions pass and all four French screenshots are readable. The fixture
omitted its destination-folder field; final Debug source corrects this but is a different receipt.
See [visual review](../../../../Plans/WIP/FileOperationsUiReview_2026-09-16/implementation/2026-09-19-throughput-disclosure/README.md).
Narrow footer clipping remains open. No full visual/input or repository acceptance is claimed.

Expanded eight-redraw batch values: 66445, 66587, 66063, 66037, 66502 us (median 66445).
Collapsed: 67173, 66106, 67068, 66325, 66882 us (median 66882). The full stream is retained.
These results are unpaired against unchanged Release production and do not establish non-regression.
An earlier Release launch omitted the test flag and produced no case results; it is diagnostic only,
retained at `C:/RedSalamander.Perf/evidence/i26-ui/throughput-ui-release-tests-disabled`.

Archived runner output has trailing whitespace normalized; unchanged full logs remain in the governed local evidence folder.
