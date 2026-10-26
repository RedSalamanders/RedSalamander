# Full-path access: baseline-repeat

Focused x64 Debug Commands run: 0 passed, 3 failed, 0 skipped.
The unchanged fixture checks full French/Unicode UIA names, help, stable displayed-row IDs and visible
screen bounds at 480/640/760/480 DIP, plus multiple streams, replacement/removal, Delete, completed
destination-file values and empty paths. Raw provider checks are not a screen-reader or hover test.

Receipt `1ab1db16ae15fff58afe7450e6ce34cbea8a91c39e3a901434ca5ec28b949a26` binds workspace `217176e2484c78b9f883030ddfd9d483e0bbdb98881af4459aef0ea5274377aa`.
The runner reports zero build warnings/errors and zero sandbox audit issues. The preserved first
failure is: Capture throughput-collapsed__480dip__fr-FR__144dpi.png must contain legacy text and background: hr=0x8007000D, foreground=0, background=0.. Candidate passes are not Full repository qualification.

`FileOps.Popup.PathAccess.RedrawBatchUs` measures five eight-redraw batches per requested width,
in that order, on native 144 DPI with French resources and reduced-motion dark theme. All timing
samples are retained. `metrics-compaction.json` binds the original full byte/row counts and digest
to the exact selected rows and their original line numbers (selected-row digest normalizes LF); no claim-bearing sample is removed.
Full raw data remains under `C:\RedSalamander.Perf\evidence\i26-ui\path-access-baseline-repeat`. Only trailing whitespace is normalized in runner.txt.

See the I26 checkpoint for the complete multi-run comparison and remaining performance gates.
These diagnostics do not establish Release, presented FPS, allocations, hidden/idle resources,
native mixed DPI, tooltip-hover or keyboard-copy acceptance. Both implementation plans remain ACTIVE.
