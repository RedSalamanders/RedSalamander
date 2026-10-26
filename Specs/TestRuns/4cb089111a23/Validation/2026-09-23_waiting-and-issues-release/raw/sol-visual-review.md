# Non-Issues file-operation gallery visual review

Candidate: `C:\RedSalamander.Perf\evidence\i26-ui\waiting-issues-isolated-20260923\runner-artifacts\artifacts\selftest\last_run\fileops\visual-gallery`
Baseline: `C:\RedSalamander.Perf\evidence\i26-ui\alert-adaptive-hosted-20260923\runner-artifacts\artifacts\selftest\last_run\fileops\visual-gallery`

Method: SHA-256 comparison of candidate PNGs against same-name baseline, then visual inspection of all 46 non-Issues changed/new captures in twelve contact sheets; opened narrow 117/121 and 107 flyout at native resolution, and compared 70 queue-paused with its baseline. This review covers static captured frames only.

Result: No new actionable visual overlap, clipping, or missing-control defect found in these 46 captures. The narrow 117 footer wraps to two visible lines. The 120-122 waiting cards fit source rows, destination and actions without the active graph. Running and paused cards retain visible graphs and separate actions. 107 flyout entries remain readable but very long history paths end at the popup edge, consistent with the baseline 480 flyout; this review does not qualify the separate DxUi menu clipping gate. Scenario 70 retains a large empty graph-height region while queue-paused, also present in the baseline; it is the admitted-transfer active-height policy, not a new regression.

## Reviewed candidate PNGs

- `04-queued-start-now__dark__480dip__fr-FR__display1__144dpi.png`
- `04-queued-start-now__dark__480dip__fr-FR__display2__96dpi.png`
- `04-queued-start-now__dark__760dip__fr-FR__display1__144dpi.png`
- `04-queued-start-now__dark__760dip__fr-FR__display2__96dpi.png`
- `05-queued-overlap-no-start-now__dark__480dip__fr-FR__display1__144dpi.png`
- `05-queued-overlap-no-start-now__dark__480dip__fr-FR__display2__96dpi.png`
- `05-queued-overlap-no-start-now__dark__760dip__fr-FR__display1__144dpi.png`
- `05-queued-overlap-no-start-now__dark__760dip__fr-FR__display2__96dpi.png`
- `09-copy-single-stream__dark__760dip__fr-FR__display1__144dpi.png`
- `09-copy-single-stream__dark__760dip__fr-FR__display2__96dpi.png`
- `107-destination-history__dark__480dip__fr-FR__display1__144dpi.png`
- `107-destination-history__dark__480dip__fr-FR__display2__96dpi.png`
- `107-destination-history__dark__760dip__fr-FR__display1__144dpi__flyout.png`
- `107-destination-history__dark__760dip__fr-FR__display1__144dpi.png`
- `107-destination-history__dark__760dip__fr-FR__display2__96dpi__flyout.png`
- `107-destination-history__dark__760dip__fr-FR__display2__96dpi.png`
- `117-expanded-wrapped-summary__dark__480dip__fr-FR__display1__144dpi.png`
- `117-expanded-wrapped-summary__dark__480dip__fr-FR__display2__96dpi.png`
- `117-expanded-wrapped-summary__dark__760dip__fr-FR__display1__144dpi.png`
- `117-expanded-wrapped-summary__dark__760dip__fr-FR__display2__96dpi.png`
- `120-waiting-unknown-total__dark__480dip__fr-FR__display1__144dpi.png`
- `120-waiting-unknown-total__dark__480dip__fr-FR__display2__96dpi.png`
- `120-waiting-unknown-total__dark__760dip__fr-FR__display1__144dpi.png`
- `120-waiting-unknown-total__dark__760dip__fr-FR__display2__96dpi.png`
- `121-waiting-move-auto__dark__480dip__fr-FR__display1__144dpi.png`
- `121-waiting-move-auto__dark__480dip__fr-FR__display2__96dpi.png`
- `121-waiting-move-auto__dark__760dip__fr-FR__display1__144dpi.png`
- `121-waiting-move-auto__dark__760dip__fr-FR__display2__96dpi.png`
- `122-waiting-retained-rows__dark__480dip__fr-FR__display1__144dpi.png`
- `122-waiting-retained-rows__dark__480dip__fr-FR__display2__96dpi.png`
- `122-waiting-retained-rows__dark__760dip__fr-FR__display1__144dpi.png`
- `122-waiting-retained-rows__dark__760dip__fr-FR__display2__96dpi.png`
- `14-paused__dark__760dip__fr-FR__display1__144dpi.png`
- `14-paused__dark__760dip__fr-FR__display2__96dpi.png`
- `30-running-and-queued__dark__480dip__fr-FR__display1__144dpi.png`
- `30-running-and-queued__dark__480dip__fr-FR__display2__96dpi.png`
- `30-running-and-queued__dark__760dip__fr-FR__display1__144dpi.png`
- `30-running-and-queued__dark__760dip__fr-FR__display2__96dpi.png`
- `70-queue-paused__dark__480dip__fr-FR__display1__144dpi.png`
- `70-queue-paused__dark__480dip__fr-FR__display2__96dpi.png`
- `70-queue-paused__dark__760dip__fr-FR__display1__144dpi.png`
- `70-queue-paused__dark__760dip__fr-FR__display2__96dpi.png`
- `71-waiting-for-other-task__dark__480dip__fr-FR__display1__144dpi.png`
- `71-waiting-for-other-task__dark__480dip__fr-FR__display2__96dpi.png`
- `71-waiting-for-other-task__dark__760dip__fr-FR__display1__144dpi.png`
- `71-waiting-for-other-task__dark__760dip__fr-FR__display2__96dpi.png`
