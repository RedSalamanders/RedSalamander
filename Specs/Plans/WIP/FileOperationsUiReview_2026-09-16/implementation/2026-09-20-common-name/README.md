# One conflict filename and one waiting state

Implemented 2026-09-20, bounded R1 consumer slice. [Before/after gallery](gallery.html), [capture provenance](capture-manifest.json), [paired measurements](paired-measurements.json).

For ordinary name collisions with identical admitted leaves, the real popup now draws one full-width bold filename heading, followed by source and destination parent folders and their separate metadata. The question no longer interpolates the filename. The expanded task title identifies the operation; the redundant fixed-width state pill is removed. One measured waiting/loading line and the warning icon/stripe remain. Collapsed cards retain their status-bearing title.

The single heading has `FileOperations.Task.<id>.Conflict.Name` identity. The actual UIA tree contains exactly one filename mention and one waiting-state mention in the common-name fixture, including enclosing tooltip regions. Full original paths remain in the immutable prompt facts. Differing leaves, literal backslashes in a slash namespace, and unknown-profile exact-path fallback remain visible. No provider I/O, admission, conflict policy/default or shared DxUi API changes are introduced.

- [baseline: 0 passed / 3 failed / 0 skipped](../../../../../TestRuns/4cb089111a23/Commands/2026-09-20_121637-i26-common-name-baseline/README.md)
- [candidate: 3 passed / 0 failed / 0 skipped](../../../../../TestRuns/4cb089111a23/Commands/2026-09-20_122120-i26-common-name-candidate/README.md)
- [final: 34 passed / 2 failed / 0 skipped](../../../../../TestRuns/4cb089111a23/Commands/2026-09-20_122455-i26-common-name-final/README.md)

All builds report zero warnings/errors and every run has zero sandbox audit issues. The baseline precedes renderer changes and fails the heading identity assertion in all three repetitions. The candidate passes all three repetitions. The final regression run is **34/36**, not a pass: throughput repetitions 2 and 3 fail before clipboard mutation because advertised format 49343 (`PNG`) cannot be retained (`CLIPBRD_E_BAD_DATA`, `0x800401D3`). The first repetition passes; the cause is not yet established. Original failure receipts are retained. No clipboard clearing or unchanged retry is performed; byte-for-byte restoration is not established by the failed run. The other conflict, prompt lifecycle, metadata, provider, medium-contract, presentation/high-contrast, failure-surface, completed-navigation and popup-resource checks pass. This clipboard gate remains blocking.

The 11 published French harness images (three baseline, eight candidate) are directly inspected at native 144 DPI. Common names are prominent and unquoted; parent folders and metadata are complete. Median eight-redraw batches change from 66,545 to 66,573 microseconds (+0.042%) for common names and 66,568 to 66,590 (+0.033%) for differing names; this does not demonstrate a regression. These Debug fixture timings are not equal-work Release FPS or final paired resource qualification.

Remaining gates are explicit: narrow buttons still clip, one-item/differing-name hierarchy polish remains, inspection/copy/return actions are not implemented by this slice, and short viewports/collapsed headers still need measured layout adoption. Required qualified DxUi exact-main CI/pin adoption, Release/Full, mixed-DPI keyboard/AT and final paired resource validation remain open. The plans remain ACTIVE. The historical September 20 audit is preserved as before-state evidence.

## Regenerated native catalog

[Gallery receipt](../../../../../TestRuns/4cb089111a23/FileOps/2026-09-20_123047-i26-common-name-gallery/README.md): **1/1**, zero skips, warnings/errors and audit issues. All **300** French native harness frames are retained with per-image source/build/hash provenance. The ordinary 96-DPI conflict, 144/96-DPI long-name conflict and 96-DPI metadata-loading frame are directly inspected: their common heading and parent folders are complete. Narrow action clipping remains visible on both displays. The other 296 catalog frames are captured and indexed, not newly visually accepted. Together with the 11 focused frames, this folder contains 311 captures, 15 directly inspected. This is static native mixed-DPI evidence; same-HWND focus/keyboard transition and final redesigned UI acceptance remain open.

## Clipboard investigation follow-up

[Owned PNG stream characterization](../../../../../TestRuns/4cb089111a23/Commands/2026-09-20_124438-i26-clipboard-png-characterization/README.md) passes 3/3 after invalidating the original backing and repeating snapshots at EOF. It uses no user clipboard access. This rules out that direct owned-data-object boundary as a reproduction; actual OLE flush/readback and restoration remain unqualified. The 34/36 failure is not waived or replaced by these passing tests.

## Clipboard persistence repair

The [new broader matrix](../../../../../TestRuns/4cb089111a23/Commands/2026-09-20_133432-i26-clipboard-ole-final/README.md) passes 36/36 after a retained 0/3 OLE baseline, 0/1 diagnostic and 6/6 candidate. Stream-backed formats are materialized as owned native bytes, and all seven admitted user-clipboard formats are verified exactly after each directed run. The historical 34/36 failure remains above. No product pixels changed in this harness-only repair; native captures remain in the raw runs. Unsupported clipboard carriers fail before mutation. This does not close the remaining redesign, Release/Full or mixed-DPI/AT/resource gates.
