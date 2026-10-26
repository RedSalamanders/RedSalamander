# Verification progress — 2026-09-21

Actual French application captures from the unchanged non-activating
`FileOps_VisualGallery` fixture, dark theme, 480 DIP. The transfer is complete
and verification is halfway through; the task and footer now both show 75%.
Actual transferred bytes remain 8/8 GB. Verification never fabricates additional
transferred bytes. The taskbar uses the same weighted work model.

| Native display | Before | After |
|---|---|---|
| 144 DPI | [full footer too early](../2026-09-21-path-label-fit/candidate--17-verifying__dark__480dip__fr-FR__display1__144dpi.png) | [75% footer](17-verifying__dark__480dip__fr-FR__display1__144dpi.png) |
| 96 DPI | [full footer too early](../2026-09-21-path-label-fit/candidate--17-verifying__dark__480dip__fr-FR__display2__96dpi.png) | [75% footer](17-verifying__dark__480dip__fr-FR__display2__96dpi.png) |

[Build/scenario provenance, semantic contracts and paired redraw evidence](../../../../../TestRuns/4cb089111a23/FileOps/2026-09-21-verification-progress/README.md).
The native gallery passes 1/1; focused progress and finished-task contracts pass
2/2, all zero skips. These screenshots establish the corrected visible fraction;
they do not close broader keyboard/UIA, Release or performance/resource gates.
