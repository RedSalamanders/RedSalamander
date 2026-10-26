# Throughput external reload and process restart — focused Debug receipt

Source base `8aef3b41a733bd43c1442c05ee739ff47929e310` plus the test fixture and
test-only owned-popup teardown changes. Product behavior is unchanged. Qualified
DxUi pin: `78b3de389a189c7f86f611787e0489fb6d474218`.
Full test-enabled x64 Debug build: zero warnings/errors, receipt
`9b380865cd8dfa0106bfc2a89b7c98e5868cbc0374c908391785422aa928dd27`.

Runner `20260921T180233Z-38192-d88dc88b193e4328a9ec4002b0e6b4fc` passes
3/3 Commands cases, zero skipped:

| Case | Duration | Coverage |
|---|---:|---|
| `cmd_pane_fileops_popup_throughput_disclosure` | 15,283 ms | Existing disclosure, suppression, geometry and failed-save recovery assertions |
| `cmd_pane_fileops_popup_throughput_external_reload` | 1,144 ms | External disk saves reach the main-window watcher/merge/apply route; existing card states survive and new cards adopt the preference |
| `cmd_pane_fileops_popup_throughput_restart` | 14,020 ms | Three owned application processes share isolated settings; startup and shutdown retain choice and both placement records |

The child modes are default-expanded then collapse, restarted-collapsed then
expand, and restarted-expanded. Each child exits successfully; the parent reads
settings after exit and compares position, dimensions, state and DPI for both
`FileOperationsPopup` and `FileOperationsPopupExpanded`. Raw child result files
and process/exit receipts are retained in `throughput-restart/`.

External collapsed/expanded settings apply in 302,794/293,271 microseconds.
Queued save generation stays 25→25 in both cases: reload and painting cause no
new settings write. `throughput-metrics-subset.jsonl` is a derived selection of
the two `FileOps.Popup.Throughput.ExternalReloadUs` rows, with original field
values. This characterizes unchanged behavior; it is not a paired overall
performance/non-regression claim.

The aggregate and Commands result JSON and child artifacts are byte copies.
Full logs and unabridged artifacts remain under
`C:/RedSalamander.Perf/evidence/i26-ui/throughput-persistence-combined-v2-20260921/last_run/`.
The full `perf/perf_metrics.jsonl` exceeds the per-file archive cap; its SHA256 is
`fc7473eecfefa856cd8ebc5cdcaf44f14656ff20eca5b122a681d8f27303a5b3`.
Build log: `C:/RedSalamander.Perf/evidence/i26-ui/throughput-persistence-final-build-20260921.log`.
Runner log: `C:/RedSalamander.Perf/evidence/i26-ui/throughput-persistence-combined-v2-20260921.log`.

The runner reports 260 disk-audit findings; no clean-disk claim is made. No new
restart screenshot is claimed. The unchanged disclosure fixture regenerates
its existing French harness images in the full artifact copy. These focused
Debug passes do not establish final Full/Release, native mixed-DPI/AT or paired
resource acceptance.
