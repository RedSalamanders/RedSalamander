# Default File Operations family dispatch

The broad Release runner exposed a test-registration defect: the new exact Batch Rename phase passed when selected directly but was missing from the default child-family table. V5's application process completed normally and reported the case as an unexpected `not executed` skip. The governed wrapper correctly rejected it with `FORBIDDEN_SKIP_CLASS, UNALLOWED_SKIP` and exit 1. This failed run remains unchanged.

V6 registers the dedicated Batch Rename family and adds a compile-time invariant requiring every default phase except setup/cleanup to belong to a dispatched family. Existing family overlap remains valid. Only 26 lines in the test registration translation unit change; production and all other 35 frozen source inputs remain unchanged. The source, patch, frozen input manifests and full-solution build receipts are retained. Both native profiles build with zero warnings/errors.

| Run | Native passed / failed / skipped | Wrapper exit | Copied files independently rehashed |
| --- | ---: | ---: | ---: |
| v5-release | 194 / 0 / 22 | 1 | 34 |
| v6-release | 195 / 0 / 21 | 0 | 41 |
| v6-debug | 196 / 0 / 20 | 0 | 41 |
| v6-debug-common | 1 / 0 / 0 | 0 | 27 |

Both V6 broad runs execute the Batch Rename case successfully through default family dispatch. Each also executes the real Recycle Skip carrier case. Debug additionally runs the opt-in common presentation fixture; this is a runtime smoke, not a paired timing result. All four disk audits have zero findings. Release's declared skips cover its Debug-only scheduler export, two 7z scenarios owned by Compare coverage, and absent remote profiles; Debug's exact skips are retained in the analysis and native receipts. Declared skips are not treated as executed capability passes.

The Debug run was strengthened from three focused cases to broad FileOps plus the common smoke before its build started. Its timestamped decision and driver are retained. A single sequential lane owns all builds/tests, uses the governed nonactivating runner, and explicitly authorizes the marked alternate-volume test root. No foreground interaction or screenshot is performed by this change.

The isolated consumer remains on base `ccebd5af68eb45f2e9deed64a7063780e71eb1ce`, with the exact 36 overlays and qualified DxUi pin `78b3de389a189c7f86f611787e0489fb6d474218`. The candidate grid/menu branches are not adopted. These results qualify the current x64 Debug/Release FileOps suite slice only. Full runner, foreground journeys, UIA/AT, final combined pin, native monitor transitions/192 DPI, final gallery and integrated resources remain separate gates. No performance exception is accepted here.

Full immutable runner trees, including complete large telemetry streams, remain beneath `C:/RedSalamander.Perf/evidence/i26-ui/` in the three directories named by the retained curator. Their manifests bind all files; the reviewed packet copies complete claim-bearing results, traces, workspace/source snapshots and runner logs without truncation. Raw JSON uses `.receipt.txt` suffixes without changing bytes. `SHA256SUMS.txt` covers every payload; `raw/file-map.receipt.txt` maps copies to their sources.
