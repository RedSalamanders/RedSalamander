# File Operations graph labels: graph-labels-resource-paired

First candidate resource comparison against resource-lifetime-measured. Timing is investigated in the subsequent A/B repeat; no performance acceptance is inferred.

Focused x64 Debug result: **3 passed, 0 failed, 0 skipped**. Build: 0 warnings, zero errors. Runner audit: zero issues. Exact workspace/build identities are retained beside this document.

The resource fixture runs five warm-up and twenty measured popup lifetimes per repetition, with 32 synchronous redraws before hiding and 32 after showing. Three repetitions cover 75 lifetimes and 4,800 explicit redraws. Actual client geometry settles at 1140 x 576 physical pixels (760 x 384 DIP at 144 DPI), despite requesting 760 x 480 DIP. Every destruction must return host/device attachment counters to baseline. Hidden intervals pump real messages for at least 350 ms and require no new hosted frames.

Private bytes, working set, process/GDI/USER handles and cumulative UI/process CPU accounting are retained for every phase. CPU accounting is quantized; a zero delta does not prove zero CPU or wakeups. These bounded lifetimes do not establish GPU residency, allocation, long-soak or Release acceptance.

`metrics-compaction.json` binds the retained metric rows to their raw source SHA-256, original line numbers and normalized-LF digest. All selected samples remain in order. Raw data and logs remain under `C:/RedSalamander.Perf/evidence/i26-ui/graph-labels-resource-paired`. Runner text only normalizes trailing whitespace.

No Fresh Full, Release, native mixed-DPI, screen-reader or final performance gate is closed by this focused run.
