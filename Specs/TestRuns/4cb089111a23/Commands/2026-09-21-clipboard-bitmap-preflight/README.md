# Bitmap preflight repair; native OLE qualification remains open

Machine `4cb089111a23`, x64 Debug, qualified DxUi `78b3de3`, consumer `eed511720`
plus the exact dirty snapshots recorded in each unmodified runner aggregate.

`preflight-failed` is 2 passed / 1 failed, zero skips: the throughput fixture
refused clipboard format 2 (CF_BITMAP) with DV_E_TYMED before any mutation.
It never reached the newly instrumented OLE flush. This failure does not identify
the cause of the earlier flush/restoration failure.

`owned-passed` is 1 passed / 0 failed / 0 skipped for
`cmd_pane_fileops_clipboard_snapshot_medium_contract`. The test deep-copies a
synthetic bitmap, replaces its source through IDataObject, proves the source
value changed, and verifies retained pixels after source destruction. Existing
stream-backing/cursor/corruption checks also run. It never publishes to the desktop
clipboard. The stable build reports zero warnings/errors.

Results, traces and source/build provenance remain intact. Full artifacts/logs are
under `C:/RedSalamander.Perf/evidence/i26-ui/clipboard-flush-diagnostic-20260921`
and `clipboard-bitmap-owned-final-20260921`. The later quiet/ownership guard is
still under review and has no accepted native verdict. Scratch-audit findings
remain in aggregates; native restoration and final qualification are not waived.
