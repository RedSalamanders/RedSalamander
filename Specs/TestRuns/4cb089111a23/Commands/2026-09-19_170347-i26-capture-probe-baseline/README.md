# Capture probes before repaint retries

Run `20260919T150347Z-27104-acc037f5d20e42e884ec5d6bf0a39aea` passes 3/3 Debug repetitions,
zero skips and zero sandbox issues. Build: zero warnings/errors. Receipt
`a5384e913aa722f6a1a03deafc7b9cf13c97c84cdf70cc668eb7beb6eca78d75` binds source
`b3a71667adf11e318f6b3b04ade7e8298b1223fcafab450dd55ac8bc179b224a`.

This adds pixel validation but retains the original first-frame-only capture path. All 12 observed
label witnesses pass (35 foreground / 4859 background pixels). This proves the previous missing
layer was intermittent; it is not a failing baseline for a deterministic renderer repair. Earlier
rejected images and their exact receipts remain in the preceding throughput slice.

Suite duration is 5810 ms for three repetitions. Keep the full metric stream; this diagnostic
observation makes no performance or root-cause claim. Runner-log trailing whitespace is normalized;
unchanged raw artifacts remain at `C:/RedSalamander.Perf/evidence/i26-ui/capture-layer-repeat`.
