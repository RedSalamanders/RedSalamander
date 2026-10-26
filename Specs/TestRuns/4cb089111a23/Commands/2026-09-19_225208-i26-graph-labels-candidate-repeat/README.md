# File Operations graph labels: graph-labels-candidate-repeat

Repeated final candidate using the same resource fixture and runner selection. All repetitions are retained; see comparison.json for unresolved timing qualification.

The [complete A/B/A/B record](comparison.json) preserves all four runs and the identical fixture
digest. Initial baseline medians are 243,928 / 244,233.5 / 243,354.5 us per 32 redraws; initial
candidate medians are 255,824.5 / 253,749 / 253,442 us. Repeated baseline medians are
242,986.5 / 244,851.5 / 243,820 us; this unchanged candidate repeats at 246,061 / 245,517.5 /
244,920.5 us. The larger slowdown did not reproduce. The remaining difference and Release
qualification remain open; no threshold is relaxed and no unfavorable run is discarded.

Warm-to-settled private-byte deltas here are +258,048 / -909,312 / -1,470,464 bytes. The repeated
baseline also has a first-round increase (+376,832 bytes). Every destruction returns host/device
attachments to baseline. These bounded observations do not establish a leak-free long-duration soak.

Focused x64 Debug result: **3 passed, 0 failed, 0 skipped**. Build: 0 warnings, zero errors. Runner audit: zero issues. Exact workspace/build identities are retained beside this document.

The resource fixture runs five warm-up and twenty measured popup lifetimes per repetition, with 32 synchronous redraws before hiding and 32 after showing. Three repetitions cover 75 lifetimes and 4,800 explicit redraws. Actual client geometry settles at 1140 x 576 physical pixels (760 x 384 DIP at 144 DPI), despite requesting 760 x 480 DIP. Every destruction must return host/device attachment counters to baseline. Hidden intervals pump real messages for at least 350 ms and require no new hosted frames.

Private bytes, working set, process/GDI/USER handles and cumulative UI/process CPU accounting are retained for every phase. CPU accounting is quantized; a zero delta does not prove zero CPU or wakeups. These bounded lifetimes do not establish GPU residency, allocation, long-soak or Release acceptance.

`metrics-compaction.json` binds the retained metric rows to their raw source SHA-256, original line numbers and normalized-LF digest. All selected samples remain in order. Raw data and logs remain under `C:/RedSalamander.Perf/evidence/i26-ui/graph-labels-candidate-repeat`. Runner text only normalizes trailing whitespace.

No Fresh Full, Release, native mixed-DPI, screen-reader or final performance gate is closed by this focused run.
