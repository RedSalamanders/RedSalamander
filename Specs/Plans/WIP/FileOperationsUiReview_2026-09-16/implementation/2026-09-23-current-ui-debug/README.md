# Current File Operations UI — x64 Debug

The frozen `a3777b3f0` consumer UI builds with zero warnings/errors and passes fresh
Debug gallery plus PR validation on the unchanged qualified DxUi pin `78b3de3`.
This validates the confirmation and compact waiting/typed Issues slices in this
configuration. It is not Full, native input/AT, performance or final pin acceptance.

- Full-solution test-enabled build: 04:13, receipt
  `e170a3830988bb762190b31c3baecf592b9d90b9d9b23dd32b087ffc57f0c8de`.
- Gallery `20260923T155151Z-65476-bd94c626f988456388e6eaac6e2236d4`: 1/0/0,
  166,838 ms runner time, disk audit zero; 424 original PNGs.
- PR `20260923T155519Z-94308-a69c3a91ee364255b73c3c6b17edec56`: 824 reported passes,
  0 failures, 0 skips; 623,228 ms,
  disk audit zero. Native assertions and Pester tests remain separately detailed in logs.

All 48 same-HWND transition checks pass at 480/760 DIP on physical 96/144-DPI
displays, with at least four intermediate/settled paints per observation. All
424 image pairs were compared to the reviewed Release originals: 100 are
byte-identical; 324 differ solely in the title-bar app icon (162 per DPI).
Every body pixel and image dimension is identical. The resource script selects
logo_debug.ico under _DEBUG. Two Sol High reviewers checked contact sheets,
originals and the comparison; root independently verified the complete pixel
bounds and representative originals. Contact sheets are analysis aids only.

The review retains existing defects: the French speed-limit Selector clips both
ends of its current value; completed-group captions and long terminal headings
also need review; the known destination-history flyout remains clipped. The
speed-limit issue is being addressed by the next measured-width slice. These
passing harness runs do not imply that every reviewed UI defect is resolved.

The isolated checkout remains based at ccebd5af with the sixteen explicitly
hashed source overlays. Both complete run trees were copied before runner
cleanup under C:/RedSalamander.Perf/evidence/i26-ui/waiting-issues-debug-20260923.
Exact build/source/runner receipts, failed-capability limits and review notes
are retained. Subsequent speed-label witness edits are outside this frozen run.

[All 424 pair hashes and pixel bounds](capture-comparison.json), [complete scenario inventory](captures/scenarios.tsv), and [gallery](gallery.html).
