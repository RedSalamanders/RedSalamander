# DxUi adoption: local Release package comparison

Both x64 Release packages pass clean-extraction application startup and built-in
plugin contracts. The baseline is master `0dd0bd6d`; the candidate is migration
`751437d2`, pinned to DxUi `13788e95`. Both use the same MSVC toolchain, vcpkg pin,
build number 100, and test-enabled configuration.

| Package footprint | Master baseline | Migration candidate | Change |
|---|---:|---:|---:|
| Compressed ZIP bytes | 30,289,465 | 30,692,818 | +403,353 (+1.33%) |
| Expanded file bytes | 78,647,558 | 79,757,185 | +1,109,627 (+1.41%) |
| Main executable bytes | 28,911,616 | 29,183,488 | +271,872 |

The candidate adds 11 provenance sidecars. Every packaged DxUi module's SHA-256
matches its sidecar and exact pin; `candidate-provenance.json` retains those values.
`package-comparison.json` includes both ZIP hashes and every changed file size.

This measures the complete migration on a test-enabled Release build. It does not
isolate linker packaging, estimate a DLL, qualify ordinary shipping-build sizes,
or measure live memory, GPU resources or timing. The earlier non-test September 9
PDB estimates describe a different build and remain separate decision context.

The unchanged master package is retained as a local rollback artifact. Its clean
extraction passes with its own harness. Neither package has been published or
installed over the user's application. A final-candidate package and consumer pin
upgrade/revert validation remain separate from this receipt.

Full logs and both ZIP files remain under the governed local directory
`Z:\RedSalamander.Perf\evidence\I19-20260909`, using the `baseline-0dd-Release-`
and `751437d2-Release-` prefixes. No hosted CI was used.
