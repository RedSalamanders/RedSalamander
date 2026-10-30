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
- [ ] Validate manifest/tool consistency, affected tooling and the actual package graph.
- [ ] Build and test the affected package consumers with exact source/dependency receipts.
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
