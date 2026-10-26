# Clipboard ownership and native contention qualification

Machine `4cb089111a23`, x64 Debug, consumer `eed511720` plus the exact dirty
snapshots in each runner aggregate, qualified DxUi pin `78b3de3` unchanged.
This is test-support qualification, not a product performance or Full pass.

| Retained stage | Passed / failed / skipped | Finding |
|---|---|---|
| owned-passed | 1 / 0 / 0 | Independent storage survives producer changes. |
| native-read-failed | 2 / 1 / 0 | Native readback failed after successful flush; original restoration passed. |
| diagnostic-passed | 3 / 0 / 0 | Instrumented confirmation, insufficient alone. |
| order-lock-failed | 8 / 1 / 0 | Repeated order identifies CLIPBRD_E_CANT_OPEN during readback. |
| read-retry-owned | 1 / 0 / 0 | Read fault injection and valid PNG/WIC seed pass. |
| read-retry-failed | 7 / 2 / 0 | Flush contention and stale screenshot probe geometry remain. |
| native-final | 9 / 0 / 0 | Three repeated shuffled rounds; all original clipboard restorations and path-focus pixel witnesses pass. |
| family-before-settlement | 49 / 1 / 0 | Clipboard passes; inspection fixture reused cached names before enumeration settled. |
| inspection-isolated | 1 / 0 / 0 | Isolated pass does not establish broad-process cleanliness. |
| family-settled | 50 / 0 / 0 | New owner settlement required before reentrant-selection test. |
| inspection-predecessor-failed | 8 / 1 / 0 | Repeat 2 selects the source but withholds focus; popup stays visible. Open investigation. |
| inspection-diagnostic-passed | 9 / 0 / 0 | All 12 instrumented handoffs preserve the input tick and transfer focus correctly. |
| canonical-commands | 912 / 0 / 2 | Canonical order passes, including native clipboard and all four inspection handoffs. |

Final qualification uses clean full Debug build receipt
`770e0bccf4ec5fa37a9be15c67aba3680b0d35245d1641742d8585e650a8c98c`.
The canonical run takes 34m 58.8s. Its two declared-capability skips are the
opt-in ViewerSpace 20k-layout and large-performance scenarios; neither is a File
Operations case. All four canonical inspection handoffs have unchanged input
ticks. The earlier guard refusal remains retained and is not attributed to a
particular outside event without evidence. Later stable-input passes do not waive
interference, mixed-DPI, assistive-technology or Full/Release gates.

The clipboard repair copies every admitted HGLOBAL/stream and GDI bitmap into
independent storage, uses valid HTML/PNG fixtures, and checks admitted owner and
generation before subsequent mutation/restoration. Read and flush retries are
bounded to three attempts and only CLIPBRD_E_CANT_OPEN; publication and product
copy are never retried. Native-final records five real read-contention recoveries,
15 explicitly injected flush-contention rows and three exact original-value
restorations. These witnesses do not claim atomic compare-and-publish, protection
from every external clipboard race, or native occurrence of flush contention in
the final run. No clipboard contents are logged.

Inspection's retained-name setup race is corrected without changing product
selection policy. Its later predecessor failure remains open: NavigationUs
records Selected with focusUnchanged=false, so the deliberate focus guard refused
handoff. Input/focus/foreground diagnostics are the next investigation, not a
reason to disable the guard or treat the repeated run as passed.

Native-final uses build receipt
`c39f7dca43cd5648b74d06fa49d177df530884d574d1b3d0200ae75f21cd92d7`.
Family-settled and inspection-predecessor-failed use
`75525f9f6a1bb0166721fcafe4774c9459f44d4c7e180988b26ffb859aa772b4`.
Both earlier full-solution Debug builds report zero warnings/errors. The final
770e0bcc build additionally qualifies the new instrumentation and changed-hunk
formatting; the two earlier receipts do not qualify those later bytes.

Results, traces and runner aggregates are unchanged. Each metrics-selection file
binds the raw metric stream by SHA-256 and byte count, lists selected line indices,
and hashes the selected rows. Full raw logs/captures remain under
`C:/RedSalamander.Perf/evidence/i26-ui/`, using the source paths in those manifests.
The native-final screenshots have separate repetition directories and were all
reviewed for the full French path and contrasting focus outline at 144 DPI.
They do not establish the complete mixed-DPI/AT acceptance matrix.

Disk audit findings remain visible in the aggregates (family-settled: 250;
canonical-commands: 253). Commands-family reports repository NOT_EVALUATED.
Canonical Commands now passes with the two declared opt-in skips above. Fresh
Full, Release, remaining UI integration, mixed-DPI/AT and paired resources stay open.

The reviewed [French keyboard-focus capture](family-settled/path-keyboard-focused__480dip__fr-FR__144dpi.png)
and its capture-provenance.json retain actual harness pixels and build identity.
build-receipts.json binds the raw build logs and their diagnostic summaries.
