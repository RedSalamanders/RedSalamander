# Readable terminal results and completed-group controls

The popup places the exact terminal result in a measured, wrapped body paragraph before actions.
The completed-group clear caption uses its actual hosted font width; headings and all five nonzero
result badges occupy whole measured rows. Cache entries are tied to current content, width and
font, and released when their content is hidden or removed. No DxUi pin changed.

The corrected baseline reproduces 28 group-caption failures, 56 summary failures and 12 same-HWND
short/long/original summary failures. The first witness lacked the actual terminal action geometry
and is retained only as a superseded attempt, not a paired performance baseline. Release and Debug
candidates pass those checks, all 140 speed-caption checks (including maximum uint64 speed), the
pure completed-group geometry tests and the existing caption/transition checks. The two builds
have zero warnings/errors. Gallery and PR receipts have zero failures and sandbox-audit findings;
the intentionally failing baseline is identified separately below.

The full 494-frame baseline/candidate comparison and all 494 Release/Debug pairs are retained.
Release/Debug dimensions and body pixels match; allowed differences are confined to the build's
title-bar icon. Root checked the independent Sol High review of all 88 changed captures and
inspected original group, summary, interrupted-Move, verification-failure and high-contrast images.
These are actual deterministic harness screenshots at native 96/144 DPI, in French, at 480/760 DIP.
The native input, UIA/AT, 192-DPI, Fresh Full and combined library-pin gates remain open.

`raw/layout-comparison.json` preserves five-round synchronous redraw/process observations paired
by exact scene, theme, width and display. Across 24 pairs, redraw median changes range from
-3.49% to +6.48%; private-byte deltas vary from -15.07 to +8.48 MiB, with unchanged GDI counts.
These sequence-wide observations do not isolate retained UI heap, demonstrate resource
non-regression, or establish presented FPS, GPU bytes or idle CPU. No memory tradeoff is accepted.

Two existing semantic limitations remain explicit: Partial and Indeterminate group badges share
an icon/color, and some completed cards repeat their Done label. The next status-presentation slice
will distinguish these states; this packet does not certify holistic UX completion. Complete raw
runs remain under `C:/RedSalamander.Perf/evidence/i26-ui/terminal-group-*-20260923`.

| Gallery profile | Run | Runner duration (ms) |
|---|---|---:|
| baseline | `20260923T170109Z-33836-a534eed3c1d34a4eaf9f1776f8c04845` | 149298 |
| release | `20260923T170945Z-95452-d5f4d35779294d40ad7017640980106c` | 152016 |
| debug | `20260923T172835Z-94064-f12b888b02ae4064801570661c65df22` | 171186 |

| PR profile | Run | Runner duration (ms) |
|---|---|---:|
| release | `20260923T171332Z-96388-659c1034cbb3434595f103379ef5f5e2` | 469876 |
| debug | `20260923T173149Z-94064-ec5db019ea0549af89d1cd0a05b72511` | 555217 |

Baseline/Release: 494 images, 88 changed.

[Gallery](gallery.html) · [Baseline/Release comparison](baseline-release-comparison.json) · [Release/Debug comparison](release-debug-comparison.json).
