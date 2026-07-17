# Portable ZIP Packaging

## Scope

This specification defines the fail-closed portable ZIP contract implemented by
`Installer/zip/build-zip.ps1` for Release and diagnostic Debug packages.

## Runtime dependency source of truth

- Root `RuntimeDependencies.props` is the single source of truth for app-local plugin runtime DLLs.
- `Directory.Build.targets` filters that manifest by project, configuration flavor, and platform, copies required
  inputs with MSBuild `<Copy>`, fails each missing required input with `<Error>`, and removes obsolete outputs.
- Plugin projects must not add consumer-local `PostBuildEvent`/`xcopy` batches for dependencies represented by the
  manifest. Add or amend a manifest item instead.
- Debug and ASan Debug use the manifest's Debug flavor; Release uses its Release flavor. Platform-specific names
  must be explicit for x64 and ARM64.
- Optional runtime files may be copied when present but must not be required by packaging. Removed/obsolete files
  are forbidden in package input and output.

## Package construction

The ZIP contains the shipping launchers and applications, root runtime DLLs, `Plugins\`, `Lang\`, `Themes\`, the
settings schema, license, and portable README. Build/link artifacts and the package-smoke test executable are not
shipped. The Microsoft Visual C++ runtime for the target architecture is bundled app-locally.

Before compression, packaging validates the build output through `Tools/RuntimeDependencies.psm1`. A missing
required runtime dependency or retained obsolete dependency aborts packaging and names the offending path.

## Clean-extraction smoke contract

Every ZIP is expanded into a new directory by `Tools/PortablePackageSmoke.psm1` before it is accepted.

- The four shipping executables and all 14 built-in plugin DLLs must be present.
- The extracted `Plugins\` directory must satisfy the same runtime-dependency manifest as the build output.
- When the host can execute the target architecture, `RedSalamander.exe --help` must enter the packaged process and
  exit successfully.
- The matching `PluginContractTests.exe` is copied into the extraction for validation only and run with
  `--package-smoke`. That mode loads every built-in plugin and validates enumeration, required exports, schemas,
  filesystem capabilities, and invalid-ID handling. Debug-only selftests and runtime-refresh probes remain part of
  the normal harness and are deliberately outside the package-smoke mode.
- ARM64 packages built on an x64 host receive full file/dependency validation; execution is deferred to an ARM64
  host. x64 packages execute on x64 and ARM64 Windows hosts.
- The extraction is removed after the check. The test harness is never added to the ZIP.

The Release workflow calls `build-zip.ps1`, so this smoke is a required release-package gate rather than an optional
post-release check.

## Tool and build reproducibility

- `vcpkg-tool.json` pins the vcpkg executable checkout separately from `vcpkg.json`'s ports baseline.
- Local dependency installation and CI must use the exact tool commit. A bundled or arbitrary PATH installation
  without a matching Git checkout is rejected.
- CI cache keys include both pin files and CI must not mutate user-wide MSBuild state with `vcpkg integrate install`.
- `build.ps1 -MaxCpuCount <N>` may bound MSBuild workers when a machine cannot reliably initialize the default
  number of compiler/resource processes. Zero retains MSBuild's default.

## Regression proof

`Tools/Tests/BuildReproducibility.Tests.ps1` owns the source and behavior contracts for the manifest, missing-file
failure, clean extraction, vcpkg identity, ARM64 gate, and ASan workflow. A real x64 package closeout must also run
the packaged app and `PluginContractTests.exe --package-smoke` from the fresh extraction.
