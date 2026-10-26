# Capture content guard — Release

Run `20260919T151219Z-26152-fef8dbf0834e4dcda73aa5fbb9197516` passes 1/1, zero skips and
zero sandbox issues. Test-enabled Release build: zero warnings/errors. Receipt
`a32989defe7d81e1061db01ddc1e23594233158993814c026c8fc4488aecc1ab` binds the same source
snapshot as final Debug: `3860bf5dceaa9fab857b1d13cc5f617efe7e24c8ebbdcac63fc5dd74a7e9212f`.

All four foreground/background witnesses pass with counts 35/4859; the absent-color negative
control correctly rejects and retains its image. All four French screenshots are visually reviewed
in the [gallery](../../../../Plans/WIP/FileOperationsUiReview_2026-09-16/implementation/2026-09-19-throughput-capture-validation/README.md).
Suite duration is 2220 ms. This is focused capture/state qualification, not Full, mixed-DPI, UIA or
paired Release rendering/resource acceptance. The original intermittent missing layer remains
unattributed; the helper now rejects missing evidence rather than accepting encoder success.

Runner-log trailing whitespace is normalized. Unchanged raw artifacts remain at
`C:/RedSalamander.Perf/evidence/i26-ui/capture-layer-candidate-release`.
