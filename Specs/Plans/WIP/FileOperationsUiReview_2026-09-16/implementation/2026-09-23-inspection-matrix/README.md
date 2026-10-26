# Inspection activation matrix

Actual French application captures from the deterministic Commands harness, not mockups.
The test now exercises source/destination in either last-active pane with UI Automation,
Space and Enter: twelve combinations. Previously role and key were tied to one pane.

Run `20260923T073719Z-42432-e1fe95a19d3b4e7cacc6659c1bce0835` used x64 Debug,
git HEAD `582691add64fb1cbf7baa80392950241ad90618e` plus the test changes recorded
by source snapshot `f7dea5f158ac5de23a7bcc4ed26c89e6c8759e2116323ec01cad17984ae16b6c`.
Build receipt: `59b68183749a7848a4bcef8b5ff3695d780ef4ff399d7404f3ea119676e7fd02`.
The application still uses qualified DxUi pin `78b3de389a189c7f86f611787e0489fb6d474218`.

## Outcome and limits

The repeated case reports **two passes, one failure, zero skips**. Overall exit is 1;
the selected E-drive sandbox audit has zero findings. In repetition one, Zoom PID 74904
took foreground focus while the input tick changed, during UIA destination inspection in
the right pane. The product refused the stale focus/hide handoff. Preserve this failure;
it does not establish uninterrupted acceptance. Both later repetitions passed all twelve
handoffs and same-popup returns, including exact item/pane focus and unchanged decision,
scope, disclosure/scroll and safe default/cancel behavior.

The twelve named return PNGs are the final successful repetition. They were individually
reviewed: the waiting message appears once, the filename is a separate bold line, complete
French actions wrap as whole buttons, and keyboard return shows the invoking focus ring.
The short body scrolls to the invoking inspection action; some comparison fields therefore
remain outside its viewport. UIA invocation does not imply keyboard focus-ring presentation.
These are short-filename interaction fixtures, not long-French-filename or mixed-DPI acceptance.
The capture batch does not record native DPI, so none is inferred from PNG dimensions.

The subsequent extension adds a named capture of the selected main pane before each return.
It builds with zero warnings/errors, receipt
`532177947fa3c3872676ec8eff4134bf7043b9194a3865ceb1d44d38d5fded0a`, snapshot
`7bd09cf8d7f0558500da29a13374e7734a5af6ccdeeed6ada0491a0834ca0b7b`.
It has not run: desktop availability is pending after the Zoom interruption. No selected-pane
PNG is invented or substituted from a different run. Fresh repeated input validation and
Full Debug on the clean E root remain required.

## Evidence

[Gallery](gallery.html), [capture manifest](manifest.json), and byte-preserved `raw`
receipts retain the failed trial alongside both passes. `sha256.receipt.txt` binds every
archived evidence file. Full external artifacts remain in
`C:/RedSalamander.Perf/evidence/i26-ui/inspection-matrix-20260923/attempt1/`.

Reproduction, after desktop availability is confirmed, uses the existing warning and lease:

```powershell
$env:RSBuildEnableTests = 'true'
./Tools/Run-AllTests.ps1 -Suite Commands -Configuration Debug -Platform x64 -SkipBuild -TestRoot E:/RedSalamander.Perf -CaseFilter cmd_pane_fileops_conflict_close_hides_escape_cancels_snapshot -SelfTestRepeat 3
```

Refresh attestation with the normal test-enabled build if the source/HEAD has changed.
