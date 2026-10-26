# Exact-path keyboard: resource comparison

The isolated resource case passes **3/3**, zero skips, with a zero-warning/error build and clean runner audit. The same digest-bound fixture as the retained graph-label candidate runs five warm-up plus twenty measured popup lifetimes per repetition: 75 lifetimes and 4,800 synchronous redraws. It records actual auto-sized geometry rather than assuming the requested size. All resource rows and every repetition are retained.

`comparison.json` records baseline/candidate medians, p95, private-byte deltas and quantized hidden UI CPU. `summary.json` retains all samples, destroyed-instance handles and CPU intervals. `metrics-compaction.json` binds unchanged selected records to raw source lines and hashes. Raw logs/data remain under `C:/RedSalamander.Perf/evidence/i26-ui/path-keyboard-resources`.

Host/device attachments return to baseline on every close; hidden intervals submit no hosted frames. This bounded witness does not establish zero allocations/wakeups, GPU residency, long soak, screen-reader, physical mixed-DPI or final Release/Full acceptance. See the I26 checkpoint for interpretation of timing variability.
