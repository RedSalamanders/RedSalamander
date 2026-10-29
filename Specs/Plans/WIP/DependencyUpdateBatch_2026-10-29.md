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
- [ ] Integrate the original PR commits without resetting unrelated worktrees.
- [ ] Review and refresh the complete execution-contract digest.
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
