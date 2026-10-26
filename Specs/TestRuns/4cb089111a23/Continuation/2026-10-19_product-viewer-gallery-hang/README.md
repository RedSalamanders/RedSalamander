# Viewer documentation fixture teardown hang

Status: original hang preserved; corrected source-bound x64 Debug gallery passes.

The exact `ViewerPETests TestViewerDocumentationGallery --no-activate` process
89384 (owned contained child of 104936) captured Text and Hex, then hit the
180,000 ms containment deadline. This was a hang, not an observed crash.

The fixture waited for the filesystem reference count to return to one while
still holding the viewer object. `ViewerText::Close()` closes its HWND; the
viewer retains its owning filesystem member until the object is released.
The correction must release viewer/information references before the quiet-point
wait, preserve the filesystem and DLL until callbacks retire, and flush manifest
rows so a later failed teardown cannot hide completed scenario evidence.

The raw source, generated files, two PNGs and minidump are preserved privately at
`D:\RedSalamander.Perf\evidence\product-scenarios-20261019\viewer-first-attempt`.
Their preservation manifest records independent ZIP readback and SHA-256. No
failed-attempt frame is published. The test-root policy takes precedence over
the crash-forensics skill's historical AppData/legacy artifact search examples;
this investigation stays inside the selected marked sandbox.

The final fixture releases both COM interfaces before waiting for provider references,
keeps shared Web/JSON/Markdown modules loaded across roles, and invokes the plugin
shutdown quiet point before unload. Native manifest rows are flushed per capture.
The repaired gallery passes all eight viewer captures in 13,910 ms; the eleven-state
Space case passes in 5,044 ms, with zero owned foreground transitions. See the
[successful gallery archive](../../Validation/2026-10-19_product-scenarios/README.md).
The private minidump is retained as collection evidence; no symbolic stack analysis
is claimed, and failed-attempt frames are not published.
