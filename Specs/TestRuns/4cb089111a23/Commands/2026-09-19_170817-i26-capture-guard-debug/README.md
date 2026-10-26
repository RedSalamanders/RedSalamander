# Capture content guard — Debug

Run `20260919T150817Z-42640-3ae9c86a7496424fa396524f225ab1fd` passes 3/3 repetitions,
zero skips and zero sandbox issues. Build: zero warnings/errors. Receipt
`28cbb9e2a7a85e2f64d3d305b32e39513f7005ea6f8b2aac204d0e811365c9f8` binds source
`3860bf5dceaa9fab857b1d13cc5f617efe7e24c8ebbdcac63fc5dd74a7e9212f`.

All 12 legacy foreground/background witnesses pass, with counts 35/4859. Each case's absent-color
negative control rejects the image with ERROR_INVALID_DATA and retains its PNG. Rejected frames
request a bounded non-activating repaint while capture is active. All four final named screenshots
are reviewed in the [gallery](../../../../Plans/WIP/FileOperationsUiReview_2026-09-16/implementation/2026-09-19-throughput-capture-validation/README.md).

Suite duration is 6149 ms for three repetitions, versus 5810 ms before the negative control and
retry path. Workloads differ; this is not a performance regression comparison. Production rendering
is unchanged. Full metrics/results retain all repetitions even though the named PNGs are overwritten.
The initial intermittent missing layer remains unattributed, not a proven renderer repair.

Runner-log trailing whitespace is normalized. Raw artifacts and the negative-control image remain at
`C:/RedSalamander.Perf/evidence/i26-ui/capture-layer-candidate-debug`.
