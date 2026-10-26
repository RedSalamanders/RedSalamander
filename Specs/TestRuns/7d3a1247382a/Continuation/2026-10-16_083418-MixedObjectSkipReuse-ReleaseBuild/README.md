# Mixed-object Skip reuse: Release build continuation

The remaining Release executable update completed successfully on 2026-10-16. The previously
blocking independently launched application was closed before the build; no application was
terminated by this operation. The earlier PID 51848 block remains historical evidence in the
[focused Debug record](../../FileOps/2026-10-15_175819-MixedObjectSkipReuse-Debug/README.md).

- Source: clean local master `f93415e814dc4ba89c6f91aaf5d4e87c612769a6`.
- Profile: x64 Release, application and bundled-runtime target `redsalamander`, tests disabled.
- Command: `build.ps1 -Configuration Release -Platform x64 -ProjectName RedSalamander`
  with process-local `RSBuildEnableTests=false`, restored after the build.
- Build: 2026-10-16 06:29:03 to 06:34:18 UTC (08:29:03 to 08:34:18 Europe/Paris);
  build receipt duration 311,828 ms, zero warnings and zero errors.
- [Immutable build receipt](build-receipt.json):
  `cb11b155b66d29a883366c3719c3cd5605072f352527274838144824634c1131`.
- [Raw compiler output](release-build-log.txt), [toolchain identity](toolchain-identity.json),
  attested [DxUi records](dxui-provenance/RedSalamander.exe.DxUi.json) and
  [raw-file digest catalog](raw-file-digests.json) are retained without normalization.

## Verification

[The verification record](verification-summary.json) confirms the live source/toolchain-compatible
receipt, all 133 attested output sizes and SHA-256 digests, and complete declared runtime-closure
membership. The application's closure contains 125 members. The new application digest is
`4d13a64c122273e92bc1aab9ce1ae22d9af609672435e1e5d3d55944a00d80da`.

Verification reused `Read-RSBuildReceipt`, `Test-RSBuildReceiptCompatible` and the existing
read-only `Get-RSVerifiedBuildReceiptArtifacts` helper. The tests-only receipt admission helper
was not used for this tests-disabled targeted build. All ten attested DxUi module sidecars match
the clean restored public source `e0260dde04695f0794eed79708030798b7575e0c`, API revision 4,
public-header tree, x64/Release profile, resolved fingerprint, actual archive and adjacent module
digests. The ordinary build's provenance publication already verified actual linker inputs.

## Scope

This completes the Release build that was previously blocked. It supplies build and artifact
verification, not a Release runtime or Full-repository test pass. The existing focused Debug
regression remains the runtime evidence for checked Skip across mixed file/folder/missing-metadata
shapes, with unchecked/error/profile/root isolation and unchanged Keep Both/metadata-consent
boundaries. No production source changed during this continuation. The app was not launched and
no desktop interaction, new performance comparison, ARM64 or remote-provider qualification is
claimed.
