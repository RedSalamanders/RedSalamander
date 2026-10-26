# Full-path access: candidate-repeat

Focused x64 Debug Commands run: 3 passed, 0 failed, 0 skipped.
The unchanged fixture checks full French/Unicode UIA names, help, stable displayed-row IDs and visible
screen bounds at 480/640/760/480 DIP, plus multiple streams, replacement/removal, Delete, completed
destination-file values and empty paths. Raw provider checks are not a screen-reader or hover test.

Receipt `d6910257339812ab48665e54a7d14a8d9fb9dacfaa987aefe0c70cce59d4efb8` binds workspace `ab16b687ae661c332ae2bee5a553c18d75f19a7a410d12aecc1cd1f121089c9d`.
The runner reports zero build warnings/errors and zero sandbox audit issues. All declared repetitions passed. Candidate passes are not Full repository qualification.

`FileOps.Popup.PathAccess.RedrawBatchUs` measures five eight-redraw batches per requested width,
in that order, on native 144 DPI with French resources and reduced-motion dark theme. All timing
samples are retained. `metrics-compaction.json` binds the original full byte/row counts and digest
to the exact selected rows and their original line numbers (selected-row digest normalizes LF); no claim-bearing sample is removed.
Full raw data remains under `C:\RedSalamander.Perf\evidence\i26-ui\path-access-candidate-repeat`. Only trailing whitespace is normalized in runner.txt.

See the I26 checkpoint for the complete multi-run comparison and remaining performance gates.
These diagnostics do not establish Release, presented FPS, allocations, hidden/idle resources,
native mixed DPI, tooltip-hover or keyboard-copy acceptance. Both implementation plans remain ACTIVE.

After the tests, clang-format normalized one line ending in the test fixture. The text after
normalizing CRLF/LF is identical; production bytes are unchanged from this receipt. The exact
tested/final fixture hashes are retained in [format-normalization.json](format-normalization.json).
All run-to-run comparisons use the identical pre-normalization fixture. No behavioral test rerun
is claimed for the line-ending-only cleanup.
