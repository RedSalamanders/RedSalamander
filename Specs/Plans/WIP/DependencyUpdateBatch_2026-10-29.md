# I54 Pending dependency integration — 2026-10-29

Status: ACTIVE. Owner: dependency integration. The developer requested merging all
projects and explicitly included the outstanding Dependabot updates.

This batch integrates RedSalamander PRs #130, #137, #138, #139 and #140 with the
already reviewed fold-on-main tooling in PR #150. Preserve both public history and
the history-free DualTail migration: private commit
`8b9e36191835cc71b5ee0dfeea4dddd5d942dd5a` must never become an ancestor.

## Intended package graph

- Keep the existing libjpeg-turbo minimum `3.2.0#1`; PR #130's requested value is
  already present in current main.
- Adopt 7zip `26.5`, curl `8.23.0` and SQLite `3.54.1`.
- Keep the explicit audited AWS SDK `1.11.880` and CRT `0.43.4` overrides.
- Use official vcpkg commit `aeb162b71d47950344248a4ebcac733e12f8a992`
  consistently in the manifest baseline and tool lock. The requested October 16
  baseline lacks the requested 7zip and SQLite versions. The selected October 22
  port update is the first snapshot containing all three requested versions.
- Preserve every existing dependency feature and all build, test, warning,
  provenance, cancellation and deterministic-resource gates.

Exact package versions belong in `vcpkg.json` and `vcpkg-tool.json`; the owning
build and validation contracts remain
[Build_Toolchain.md](../../Build/Build_Toolchain.md) and
[Testing_ValidationEvidence.md](../../Testing/Testing_ValidationEvidence.md).

## Checklist

- [x] Inspect every original PR patch and verify that none imports private history.
- [x] Verify the requested versions against the official vcpkg registry snapshots.
- [x] Integrate the original PR commits without resetting unrelated worktrees.
- [x] Review and refresh the complete execution-contract digest.
- [x] Validate manifest/tool consistency, affected tooling and the actual package graph.
- [x] Build and test the affected package consumers with exact source/dependency receipts.
- [ ] Pass candidate CI, merge the coherent batch and verify main.
- [ ] Publish the exact results and remaining separately owned holds; close this plan
  only when its required gates have passed.

## Evidence boundaries

Earlier qualification at the September 18 vcpkg snapshot does not qualify this
new dependency graph. Keep failed and interrupted results intact. Cross-built
ARM64 is build evidence; native runtime claims require native execution. A CI job
that has not started or finished remains pending. Do not reduce tests or change
acceptance thresholds to make this batch pass.

## Integration checkpoint — 30 October

The local integration at `092f2d00b362eac1fc6937195fd2a048dd117e09` retains
all five original Dependabot heads and the fold-on-main PR #150 head as ancestors.
It preserves main `632dc13b224715fc3e39b6b3145452495553be19`, including its
MSVC warning/ICF policy and exact DxUi pin. The package homepage now names the
public RedSalamanders repository. This checkpoint establishes integration only;
the installer, candidate CI and native consumer gates above are still pending.

Overnight qualification at the earlier PR #146 merge `28e70e6` reported viewer
desktop-restoration failures on x64, and Commands, independent-volume File
Operations and viewer failures on ARM64. Its shuffle lane timed out. These remain
failed receipts owned by I52; they are neither evidence for this new package graph
nor successful retries. Main `632dc13` passed the x64 Debug/Release and ARM64
Release PR suites; its ARM64 Debug job exceeded GitHub's six-hour limit during
the build. This batch must obtain its own actual ARM64 build and test result.

The canonical digest after reviewing the tool/registry lock change is
`e6325afc4573acc73e84f4425dce81a9a8e5bc591b75e8b3b1df4504c6ef7501`.
Four focused tooling files passed 110 checks, with zero failures and two explicit
capability skips (the canonical local vcpkg-tool path and the not-yet-restored
new DxUi pin). An independent check of the exact `aeb162b71d47` checkout found
every manifest floor and AWS override in its version database, and its pinned
tool bootstrap passed. The first local installer attempt failed while Git tried
to lazily fetch an AWS port blob: GitHub DNS resolution failed. Keep that failed
log; it is not a successful dependency install. Fresh package installation and
native candidate CI remain required.

## ARM64 overlay correction

Source `50f219ba2bf2c7cde33e252c6a515a31bdf7e5fe` subsequently passed the full
local x64 Debug, Release and ASan Debug rebuilds with zero warnings and errors.
SQLite noninteractive and Curl parser/planning checks passed in all three
profiles at source snapshot `74326aa7b222708f1dcb4ccf778e65c1cdac2ed64f80caf4cb6cae40cca5a95c`.
The intentional heap-overflow probe produced the required sanitizer diagnostic.
The two earlier local capability skips were covered by 14 passing pin/toolset
checks after the exact tool and DxUi sources were restored.

The additional ARM64 restore exposed an incompatible overlay before compilation:
the inherited `aws-c-common` 1.0.0 overlay cannot supply
`aws_last_error_or_unknown_error` and
`aws_error_or_last_error_or_unknown_error` for `aws-c-io` 1.1.0. Its failed
inputs, linker log and ABI metadata were preserved before retry. The original
hosted installer bypassed both the compatibility port and compiler-pinned triplet,
so its successful native restore did not qualify this local path.

The overlay now follows the selected registry's 1.0.3 source, archive hash, CRT
patch and notices. The exact compiler/architecture/file-scoped `/Ob1` workaround
is retained. Hosted installation uses the same triplet and ARM64 port overlays;
its cache identity includes their helper and overlay bytes. Fresh installation,
focused parity tests and candidate CI are required after this correction; the
earlier source's receipts remain source-specific.

The correction passes all 115 focused tooling checks with zero failures, skips
or pending cases. The new coverage executes the actual hosted install script's
argument selection for x64, ARM64, a failing native install and an unsupported
triplet; it also checks registry/overlay source parity and cache dependency
closure. The existing six compiler/architecture CMake controls remain passing.

## Corrected source and ARM64 build follow-up

Corrected source `e01f711f7a0852a302464388754c15b40d2c4be3` passed fresh x64
Debug, Release and ASan Debug solution rebuilds with zero warnings and errors.
Its exact source snapshot is
`1dbad1732a35b389d6061cb25022dd4a550d90c2986b8ba38b44c51571a073f8`.
The build receipt IDs are respectively
`fe74eb1341d001af4da7105f18a9c13630313c27874aabeeb9b3eb3c3fc3767b`,
`6d560aac0458445e8df71b14fd402cfdb3ab20aa2c1621f4696c1569ddc8740e`,
and `2dec30bccf98c11c1734178c15e7614fb30b4f2e0434b37612b1ebc6159948cc`.
The nonactivating SQLite group and Curl parser/request-planning checks passed
against each receipt. These are focused consumer checks, not Full or live-network
qualification. ARM64 installation also passed with the corrected overlay; the
fresh ARM64 ASan Debug cross-build passed with zero warnings and errors under
receipt `3fe8187678bd341558faca9be1fa678e5cdff0860b358ed92142c97b10b4c022`.
That receipt establishes compilation only, not ARM64 native runtime acceptance.

Local integration `d42b6a7ed3e53207a1d1e8c4b51933f339f014f5` also retains
PR #151 source `630d3565547b77268ab6bc1e3bcde5aee187471f`: native Zig selection
on ARM64 hosts and bounded contained runtime-builder children. Its reviewed
rollback contract already requires eight notices. The initial attempt to update
the generic lifecycle test changed a sealed historical input and was rejected by
the hosted integrity guard. Restore those exact bytes; current rollback coverage
belongs in the active Ghostty suite. The five required upgrade lane manifests are
a different contract and remain unchanged. Earlier receipts above remain bound
to their original sources; they are not relabeled as combined-source qualification.

## Hosted tooling follow-up

Combined source `f531c3e02e59e2812d0f7e8e5a9330e5e156f388` passed both hosted
Debug compilation steps and fresh local x64 Debug/ASan Debug plus ARM64 ASan
cross-builds with zero warnings and errors. Its local noninteractive PrePush
remainder passed 29 toolchain checks and 256 Compare Directories cases, with zero
failures and 27 declared environment skips. The runner re-attested the complete
Debug profile after the deployment proof under receipt
`7139232642a8642c5e4815382fc6770d848085a0834c6f7d92864856ba9f54e2`.
Local PrePush remained NOT_EVALUATED/exit 2: 12 PR contracts were delegated and
four focus-taking entries remained pending locally. Five pre-existing sandbox
audit items were recorded; no complete Full or globally clean sandbox claim was made.

The developer's compilation-based integration approval permitted merging #138
at `9c99afd3b3574a19bd0fe2ccd6dc17079eab1082` while hosted PR validation
continued. Its x64 job subsequently reported 1,108 passing tooling cases and two
failures: the altered sealed test and an installer assertion that still expected
the former direct command syntax/cache inputs. Retain that failed receipt; it is
not a passing candidate CI result. Main verification and this plan remain open.

The follow-up restores `TerminalDependencyLifecycle.Tests.ps1` to its recorded
SHA-256 `11154a9e66a8bb26a94795e00cf716669d00e5690db9a232f1f8b71fbff97d1e`
without changing the historical manifest. The active Ghostty test instead checks
all eight locked notice paths and their documented atomic rollback unit. The
current installer contract checks its actual argument array, pinned executable
and all five cache inputs; the existing executable argument/overlay/failure
controls remain mandatory. Validate the complete active tooling profile and
fresh hosted compilation before merging this correction, then confirm main.
