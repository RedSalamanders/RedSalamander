# Portable ZIP Packaging

## Scope

This specification defines the fail-closed portable ZIP contract implemented by
`Installer/zip/build-zip.ps1` for Release and diagnostic Debug packages.

## Runtime dependency source of truth

- Root `RuntimeDependencies.props` is the single source of truth for app-local product and plugin runtime files.
  Sources default to the pinned vcpkg tree; explicitly marked repository sources are resolved beneath the repository
  root. `OutputRoot` defaults to `Plugins` and may select `BuildOutput` for an application sibling such as the
  SearchService `sqlite3.dll`; `OutputName` is a validated relative descendant of the selected root.
- `Directory.Build.targets` filters that manifest by project, configuration flavor, and platform, copies required
  inputs with MSBuild `<Copy>`, fails each missing required input with `<Error>`, and removes obsolete outputs.
- Product and plugin projects must not add consumer-local `PostBuildEvent`/`xcopy` batches for dependencies represented
  by the manifest. Add or amend a manifest item instead.
- Debug and ASan Debug use the manifest's Debug flavor; Release uses its Release flavor. Platform-specific names
  must be explicit for x64 and ARM64.
- Optional runtime files may be copied when present but must not be required by packaging. Removed/obsolete files
  are forbidden in package input and output.

## Package construction

The ZIP contains the shipping launchers and applications, root runtime DLLs, `Plugins\`, `Lang\`, `Themes\`, the
settings schema, license, and portable README. Build/link artifacts and the package-smoke test executable are not
shipped. The Microsoft Visual C++ runtime for the target architecture is bundled app-locally.

Before compression, packaging validates the build output through `Tools/Modules/Packaging/RuntimeDependencies.psm1`. A missing
required runtime dependency or retained obsolete dependency aborts packaging and names the offending path.

The standalone ZIP entrypoint holds the selected artifact-profile lease before
reading build output, then acquires the repository-wide packaging coordination
scope before publishing beneath `.build\AppPackages`. It holds both through
compression, clean-extraction smoke, and the final archive read, and releases
packaging before the profile lease. Calls nested under `build.ps1 -Zip` reenter
the same locks on the owning thread. An abandoned packaging owner records
repository-wide contamination and fails closed with the marker path; packaging
never clears that marker automatically.

Before acquiring locks, the ZIP entrypoint rejects any already-existing reparse
component at the repository, `.build`, or `.build\AppPackages` boundary so lock or
package output cannot be redirected. Under the packaging lease it creates missing
output directories component-by-component through the guarded repository-path
helper. Compression writes a unique same-directory staged ZIP; clean-extraction
smoke validates that staged archive. Immediately before publication, the entrypoint
revalidates the output directory, staged file, final path, and every ancestor, then
uses guarded atomic move/replace. A failure before or during publication preserves
the previous complete final ZIP, and final-file reparse points are rejected.

A standalone ZIP build requires `-BuildNumber 1..65535`. It uses the pure
`New-RSVersionContext` operation for the explicitly requested configuration and
platform; it does not reuse a saved profile object or allocate a package-time
counter. `build.ps1 -Zip` passes the number that stamped the compiled artifacts.

## Clean-extraction smoke contract

Every ZIP is expanded into a new directory by `Tools/Modules/Packaging/PortablePackageSmoke.psm1` before it is accepted.

- The four shipping executables and all 15 built-in plugin DLLs must be present.
- `Plugins\TerminalRuntime\` must contain the private Ghostty DLL, its exact build identity, and the complete
  runtime notice set for Ghostty, uucode, the uucode Bjoern Hoehrmann component, Unicode data, Wuffs, and Zig. The
  nested output paths are validated as relative descendants of `Plugins\`; rooted and traversal paths are rejected.
- The packaged Ghostty identity must be byte-identical to the receipt-bound
  `Plugins\TerminalRuntime` identity from the shipping profile. Its format,
  schema, admitted commit, machine, export count/hash, and ordered import set
  must match the canonical lock. The packaged DLL hash and size must match both
  that identity and the shipping-profile DLL, and every packaged notice must
  match both its canonical normalized hash and shipping-profile copy.
- When the host can execute the target architecture, smoke loads and frees the
  packaged `ghostty-vt.dll` with the extracted profile root available for its
  app-local runtime dependencies. ARM64 defers this native check only on a
  non-ARM64 host; the native admitted-product workflow must execute it.
- The extracted profile root and `Plugins\` directory must satisfy the same runtime-dependency manifest as the build output.
- When the host can execute the target architecture, `RedSalamander.exe --help` must enter the packaged process and
  exit successfully.
- The matching `PluginContractTests.exe` is copied into the extraction for validation only and run with
  `--package-smoke`. That mode loads every built-in plugin and validates enumeration, required exports, schemas,
  filesystem capabilities, and invalid-ID handling. Debug-only selftests and runtime-refresh probes remain part of
  the normal harness and are deliberately outside the package-smoke mode. Provider mutation proofs that acquire
  TestSandbox scratch are also outside that mode and must report an explicit skip: the harness runs from a
  temporary extraction with no repository root, so sandbox acquisition fails closed there by design.
- `PluginContractTests` is the sole test project that remains an application in a tests-disabled production
  Release build. It is compiled without `ENABLE_TESTS`, attested in the same full-profile build receipt as the
  package inputs, required before ZIP construction, and never copied into the published archive. Its disabled-test
  clean path, like every disabled utility test node, removes only project-named outputs; no disabled test may replay
  stale app-local dependency entries from an earlier test-enabled `CppClean` log and remove DLLs owned by the
  shipping profile.
- ARM64 packages built on an x64 host receive full file/dependency validation; execution is deferred to an ARM64
  host. x64 packages execute on x64 and ARM64 Windows hosts.
- The extraction is removed after the check. The test harness is never added to the ZIP.
- Smoke executables run through the shared kill-on-close contained-process
  helper. Interrupting the packaging host terminates its contained smoke
  descendants before the packaging lease is abandoned and marked contaminated.

The Release workflow calls `build-zip.ps1`, so this smoke is a required release-package gate rather than an optional
post-release check.

## Tool and build reproducibility

- `vcpkg-tool.json` pins the vcpkg executable checkout separately from `vcpkg.json`'s ports baseline.
- Local dependency installation and CI must use the exact tool commit. A bundled or arbitrary PATH installation
  without a matching Git checkout is rejected.
- CI must compare the pinned commits' `scripts/vcpkg-tools.json` schema versions
  before bootstrap and require the executable schema to be at least the ports
  baseline schema. A baseline bump that raises the schema must update the tool
  pin in the same reviewed change.
- CI cache keys include both pin files and CI must not mutate user-wide MSBuild state with `vcpkg integrate install`.
  The reusable build imports `vcpkg.props` and `vcpkg.targets` directly from the pinned checkout through MSBuild's
  per-invocation `ForceImportBeforeCppTargets` / `ForceImportAfterCppTargets` properties, which must be set before
  invoking `build.ps1` so automatic linking consumes the already-installed triplet libraries. The invocation disables
  local-app-data discovery so a runner-level or user-wide integration cannot add a second, unpinned target import. It
  also sets `VcpkgXUseBuiltInApplocalDeps=false`: the pinned PowerShell copier writes copied destination paths to the
  vcpkg tlog consumed by build-receipt publication, while the built-in copier writes package source paths that the
  fail-closed receipt boundary must reject.
- `build.ps1 -MaxCpuCount <N>` may bound MSBuild workers when a machine cannot reliably initialize the default
  number of compiler/resource processes. Zero retains MSBuild's default.

## Regression proof

`Tools/Tests/BuildReproducibility.Tests.ps1` owns the source and behavior contracts for the manifest, missing-file
failure, clean extraction, vcpkg identity, ARM64 gate, and ASan workflow. A real x64 package closeout must also run
the packaged app and `PluginContractTests.exe --package-smoke` from the fresh extraction.
