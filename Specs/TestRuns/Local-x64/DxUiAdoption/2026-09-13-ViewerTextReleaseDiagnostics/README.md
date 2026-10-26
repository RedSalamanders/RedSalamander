# ViewerText diagnostics in test-enabled Release

The retained [3a43186a Release Full](../2026-09-13-Release3a43186a/README.md) fails
five Commands cases because ViewerText's snapshot handler/counters and the embedded
preview bridge require `_DEBUG`, while the cases are registered in test-enabled Release.

An isolated repeat reproduces all five failures at `3a43186a`: **0 passed, 5 failed,
0 skipped**. Correction `836f4f97` uses the existing `ENABLE_TESTS` policy for the snapshot
handler, backing counters, pointer/async diagnostics and fault hooks, and the host's preview
bridge. The same five cases then pass: **5 passed, 0 failed, 0 skipped**. Their three complete
fixture source files are byte-identical between the revisions; assertions and deadlines
are unchanged. The corrected case group takes 4,949 ms. This removes diagnostic timeouts
and does not establish a product performance improvement.

The x64 Release root build passes in 4m57s with zero warnings/errors. All 14 module
sidecars match actual output bytes, the exact DxUi `13788e95` pin and x64 compiler host.
All 33 changed preprocessor lines pass scoped formatting; whole-file formatting findings
in the untouched portions of ViewerText.Text.cpp and ViewerText.Hex.cpp match the baseline.
Specification inventory reports zero blocking findings.

`retained-files.json` hashes eleven original receipts. The comparison also records the
unchanged fixture hashes and each before/after result. Complete diagnostic archives remain
under `viewertext-release-gate-before-3a43186a` and `viewertext-release-gate-after-836f4f97`
in `Z:/RedSalamander.Perf/evidence/I19-20260909`.

No hosted CI ran. The failed Full remains failed; this focused comparison does not replace
the final x64 Full, ARM64 build/native execution, resource and package/release gates in I19.
