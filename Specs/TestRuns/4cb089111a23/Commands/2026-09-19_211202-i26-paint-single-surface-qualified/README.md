# Popup paint boundary: paint-single-surface-qualified

Focused x64 Debug run: 24 passed, 0 failed, 0 skipped.
Receipt and workspace identities, raw timing samples, captured pixel counts and first failure are
retained in summary.json, with exact source/build records alongside them. Full raw data remains at
C:\RedSalamander.Perf\evidence\i26-ui\paint-single-surface-qualified. The build reports zero warnings/errors; the runner retained one advisory for the earlier failed-build run directory. That named run was preserved under S:/RedSalamander.Perf/evidence; the final archive includes a separate clean post-archive audit. The original advisory is not rewritten.

Metrics are split only to satisfy the per-file archive cap. metrics-compaction.json binds raw byte/
row counts, SHA-256, selected line numbers and portable LF digests. No claim-bearing row is removed.
WindowState in the diagnostic baseline stores pre-draw state in durationUs, post-draw state in value0
and DPI in value1; it is a state record, not a duration measurement. CheckWindowState updates cached
state, so these diagnostic runs alone cannot establish a fix. HostedPaintUs records dispatch duration,
update-region presence and handled status; its S_OK marker is not the host's internal EndDraw result.

The covered fixture adds a disabled, non-activating owned sibling inside the popup's footprint and
captures the popup before/after removing it through the shared WGC harness. This does not take focus
or substitute a desktop screenshot. The diagnostics variant predates this extra fixture section.
All variants retain French path/UIA, disclosure, localization and capture rejection assertions.

No Release, Full, native mixed-DPI, screen-reader, hidden/idle resource or final performance gate is
closed by this focused run. See the I26 checkpoint for interpretation and the next action.
