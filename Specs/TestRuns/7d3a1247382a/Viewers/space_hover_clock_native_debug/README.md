# ViewerSpace focused native Debug evidence

Run `20261009T092014Z-27568-b30eaecd47024d9e919ec24a0fc1aa82` completed all 8 selected native cases successfully. This focused run reports repository `NOT_EVALUATED`.

Profile: x64 / Debug. Build receipt ID: `4e2c3e2030fb55d83e9dc573f1455e545d73c38729892d0f8279573427e0645f`. Source snapshot: `846115f86b923381c3d9311025957793349fecaddbccab6fa1748db82982e131`. The exact receipt is included here (95,452 bytes, SHA-256 `76b82e2dd836cf72171441ed9632832c6d3ad00fc09ae962d25683e351354722`). The launcher transcript is copied from `D:/RedSalamander.Perf/evidence/viewer-space-final-qualification-d713/native-debug-clock-fix.launch.log` (62,818 bytes, SHA-256 `f0d3124e728f7c95bcf583ebe30e74fcb6d8caf85e7997ca0d64d3c3fdd301b4`).

Logs, diagnostic JSON, captures, and original provenance were copied byte for byte; `manifest.txt` binds their source paths, lengths, and SHA-256 digests.

The separate `capture-review-sidecar.json` records root review of all nine complete frames at original resolution and acceptance for these exact lifecycle and owned-window scenarios. The eight lifecycle images retain their adjacent original provenance. The ninth image, `viewer-space-inflight-scan.png`, has no adjacent original provenance; its visual acceptance is limited to the owned-window log and immutable build receipt recorded in the sidecar.

The reviewed frames show active progress before data and during a partial scan (one folder, two files, 384 bytes); the completed nested-folder/file shape (896 bytes); retained prior results while refresh is pending, canceled, or refused; and distinct empty and provider-error states. The inflight owned-window frame has horizontal table scrolling.

This review does not qualify other locales, DPI settings, platforms, provider accounts, accessibility behavior, or performance endpoints.
