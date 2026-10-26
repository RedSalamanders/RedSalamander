# Throughput UI — final Debug state pass, rejected captures

Run `20260919T144036Z-56576-209e4ac8dfe145399bbf8508d5b1bee1` passes 1/1, zero skips,
zero disk-audit issues. Build: zero warnings/errors. Receipt
`62748c0626f03df8c63249ea3e848a85f01177dbf71743358cdf51f841afe02c`, source snapshot
`251e02931e18dcd5f88d960e8443abd1bd643bd1fafe410a33442cffe15a2317`.

State/persistence assertions pass. **Visual inspection rejects all four Debug captures:** legacy
Direct2D background/text is missing while hosted controls render. EndDraw metrics report no failed
HRESULT. Resolving dark theme/French before creation and correcting the destination fixture do not
resolve it. Keep the renderer/capture investigation open; this is not a visual pass.
See [all captures and findings](../../../../Plans/WIP/FileOperationsUiReview_2026-09-16/implementation/2026-09-19-throughput-disclosure/README.md).

Expanded batches: 376613, 377404, 373836, 373765, 374254 us (median 374254).
Collapsed: 361126, 375579, 374973, 369418, 372581 us (median 372581). These diagnostic
timings retain the unchanged measured section, but cannot qualify visible complete rendering.
Final Full, matched Release resources, save failure/hot reload, placement and input gates remain open.

Archived runner output has trailing whitespace normalized; unchanged full logs remain in the governed local evidence folder.
