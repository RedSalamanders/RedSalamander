---
name: tooling-governance
description: Add, change, move or retire RedSalamander tooling with its inventory, help, callers, compatibility, dependency closure and owning specifications.
---

# Tooling governance

Read [Tools/README](../../../Tools/README.md) and [the normative contract](../../../Specs/Testing/Testing_ToolingGovernance.md) before editing tooling. Tools/tool-inventory.json is the machine owner of classification and consumers; keep one implementation and explicit semantic exceptions.

## Classify and place the change

- Identify the command/module/test/data/compatibility/history role, direct callers, prerequisites, side effects, outputs and owning spec.
- Keep stable public entrypoints at Tools root using Verb-Noun.ps1. Put reusable implementation in explicit domain modules with reviewed exports; no import-time mutation.
- Update inventory, public help, callers, tests and owning spec in the same change. Help declares parameters, exit/output behavior, prerequisites, side effects and safe examples.
- Preserve documented external paths with thin shims where needed. Delete orphan historical implementations only after their useful evidence and caller closure are accounted for; do not edit sealed bytes.
- Use ShouldProcess for meaningful previews. Do not hide broader mutation in a narrow validation command.

## Preserve runtime and evidence boundaries

Read [build coordination](../../../Specs/Build/Build_Toolchain.md) and [selftest ownership](../../../Specs/Testing/Testing_SelfTests.md) for the resources affected.

- Artifact operations use repository plus platform/configuration. Dependency-capable phases use the narrower platform dependency lock; packaging/version/shared runtime resources retain their own short scopes and ordering. Do not serialize unrelated profiles/worktrees.
- Prove process/path/start/operation ownership. Never terminate an independently launched application. Keep automatic cleanup within the launching operation's contained descendants.
- Test paths are fail-closed beneath the selected existing marked sandbox. Parameters/environment are requests, not new-root authorization. Do not follow reparse points or expose unrelated paths/volumes.
- Only RequiresInteractiveDesktop entries use the existing warning/desktop lease, after agreement to the time. Build/setup/archival do not justify a desktop takeover.
- For affected/resume/receipt changes, read [validation evidence](../../../Specs/Testing/Testing_ValidationEvidence.md). Preserve exact source/artifact/environment/scope identity, artifact epochs, failed outcomes and explicitly pending gates; a digest refresh is not requalification.

## Keep the full dependency closure

Machine cache/receipt/workflow keys must include every actual imported module, schema, patch, notice, version and relevant environment input. Reconcile actual callers/imports with declared closure and add a focused behavioral/closure guard. Current and sealed historical manifests remain distinct.

Adding an artifact writer requires reviewed invalidation/read/write/produced sets and re-attestation barriers before readers. Restore caller environment and module ownership across nested PowerShell calls; do not force-reimport dependencies merely for convenience.

## Verify the changed behavior

Use focused behavior/negative tests for the actual command, inventory/help/caller/export contracts and dependency closure. Reuse canonical Markdown heading/link helpers in SpecInformationArchitecture.psm1 rather than another regex-only parser.

Use Test-Changes.ps1 -Explain and affected iteration, then the owning PrePush/qualification contract for final obligations. Full is explicit where required; pending CI, foreground/platform acceptance or a focused pass is never a complete repository result.

Useful checks include Get-ToolInventory.ps1 -Validate, relevant Tools/Tests suites, and Get-SpecInventory.ps1 -FailOnFindings. Run fixture-producing checks inside the governed sandbox, distinguish structural checks from task-level evaluation, and record actual outcomes. Keep the local fragment/fold/check workflow current.

## Conditional references

- [Windows discovery](references/windows-tool-discovery.md): discovery priority for developer executables.
- [Multi-target installation](references/powershell-multi-target-install.md): explicit parameter selection and target loops.
