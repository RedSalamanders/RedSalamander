# Complete ASan Full with residual failures

Source `b2cbf8f3849da29e1f6a35f53a1dc6ec0caa59e0` completed all 18
Fresh Full groups in run `20261026T022926Z-31076-da2be9551b5445949a23b0f4aa89868e`.
Sixteen groups promoted and two failed. The complete native summary records
2,557 passed, three failed and 52 skipped checks. Compare recorded 257/1/21,
Commands 954/2/6 and FileOps 297/0/25 (passed/failed/skipped).

Seven of the original eight failures now pass in this Full, including
`FileOps_SelectionShapeLocal`. The remaining original case,
`cmd_viewer_space_close_does_not_wait_for_retired_model_disposal`, measured
133,173 us against its unchanged 100,000-us limit. Its retirement submission
maximum was 69 us; destruction stages included 385 us for model retirement,
46,866 us for the renderer and 67,427 us inside `OnDestroy`, while the complete
HWND destruction measured 133,108 us. These observations do not identify the
individual renderer or nonclient resource responsible.

Two additional cases failed: the search-service stale-root seed did not finish
within its original 60-second warmup watchdog, and the dense-files observation
reported zero last-frame tile draws while logical visibility already contained
320 tiles. The related local-index cold stale-root case passed in 116,430 ms;
this duration includes both of that case's seed/query lifecycles and is not an
isolated seed measurement. No time limit, tile assertion or native verdict is
relaxed by this packet.

`full/manifest.json` binds 22 byte-identical payloads: the v2/v1 summaries,
execution plan and final ledger state, all three affected native verdicts and
their evidence/terminal records, six current-source build receipts, the failed
qualification state and a selected close-metric extract. That extract came from
the preserved Commands metric stream and was selected and UTF-8 encoded for
diagnosis; it is not represented as the original complete byte stream.

`diagnostics-manifest.json` binds the completed unchanged-binary Commands cluster:
18 checks passed across three repetitions. Those focused passes do not clear
the Full failures or prove a repair. The driver was stopped during the first
SQLite replay's preflight, with no owned native child; neither planned SQLite
replay executed. The original partial state is preserved unchanged, alongside
an explicit interruption record. The complete Full itself was not interrupted.

The Full's disk audit retains 34 existing layout findings: 28 sibling outputs
and six historical run directories. This packet makes no clean-disk claim and
does not remove historical evidence or alter the independent Release app.
The source remains unqualified for publication. Repairs, final-source validation
and the public migration remain pending.
