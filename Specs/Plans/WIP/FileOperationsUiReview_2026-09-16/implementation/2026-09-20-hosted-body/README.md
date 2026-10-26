# Retained body viewport and first-frame collapse

The accepted x64 Debug matrix passes 18/18 without compiler warnings/errors, skips or disk audit issues. Task controls share the scrolled body clip, footer controls remain separate, and Tab reveals the complete focus ring with 4 DIP clearance. The harness asserts the actual 420-DIP outer height. Checked repeat scope survives inspection and return. All three gallery frames were directly inspected.

A separate first-frame regression now removes collapsed completed-card child controls immediately while retaining the bounded height animation. The intermediate ring run failed once (17/18) because those controls previously survived during collapse; it remains archived, along with the 15/18 initial resize failure. Final (18/18) was visually rejected for its clipped focus ring; accepted (18/18) includes both corrections.

[Gallery](gallery.html) · [Provenance](capture-manifest.json) · [Paired measurements](paired-measurements.json)

The preceding keyboard baseline common/differing median/p95 was 66585/66788 and 66522/66782 us. Accepted is 66569/66783 and 66583/66708 us, 45 eight-redraw samples each. This is focused Debug evidence; old tail excursions are retained and this does not complete Release, mixed-DPI, assistive-technology or resource qualification. Required whole-action wrapping and a separately retained action region for oversized individual cards remain open.
