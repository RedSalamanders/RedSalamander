# Changes workflow

Status: normative repository workflow contract
Last reviewed: 2026-10-21

DxUi, RedSalamander, RedXe, RedPrism and RedTilde use the same fragment format, commands and behavioral fixtures.
Each repository owns its local changelog; changing one repository never updates a consumer's notes or dependency pin.

Record each change as one UTF-8 Markdown file named `Changes/<yyyy-mm-dd>-<topic>.md`, using a real calendar date
and a unique topic in lowercase letters, digits and hyphens. It contains one nonempty top-level `- ` bullet;
indent all continuation lines and sub-bullets with spaces. Describe the behavior, relevant validation and material
remaining gates. Do not maintain a second manual list or hand-edit generated `CHANGELOG.md` entries.

Keep fragments while editing so parallel work can use distinct files. **Fold before PrePush and final handoff,
and immediately after every merge/rebase that introduces fragments, even when only one is pending.** Do not wait
for a release, a large batch or a separate cleanup PR. Include the generated changelog and fragment removals in
the current change. Existing release notes remain in place; adopting this workflow does not migrate their history.

[`Tools/Test-Changelog.ps1`](../../Tools/Test-Changelog.ps1) is the read-only editing validator. It rejects invalid
names/dates, empty or multiple bullets, unindented continuation lines, unexpected entries under `Changes/`,
reparse paths, missing `Changes/README.md`, and a missing or duplicate `## Unreleased` heading in `CHANGELOG.md`.
Existing validation/tooling entrypoints invoke it or its shared module. PrePush entrypoints require zero fragments;
repositories without PrePush use the same explicit final gate below. Explanation-only runs never mutate files.

[`Tools/Fold-Changelog.ps1`](../../Tools/Fold-Changelog.ps1) validates all input before writing. It inserts entries
under Unreleased in descending ordinal filename order (newest date first, topic breaking ties), preserves the
changelog's UTF-8 BOM and line endings, atomically publishes a staged changelog, then removes only folded files.
A checkout-local exclusive lease prevents overlapping folds. Retry after interrupted cleanup does not duplicate
an identical entry already in Unreleased. An identical entry in a released section does not suppress a new entry.
`-WhatIf` previews without writing; `-Check` is read-only and rejects even one remaining fragment.
The commands never stage, commit, push, fetch, build native code or take desktop focus.
An explicit relative `-Root` resolves against the caller's PowerShell location, including after `Set-Location`,
independently of the process working directory. Validation, preview and folding all target that same checkout.

```powershell
./Tools/Test-Changelog.ps1
./Tools/Fold-Changelog.ps1 -WhatIf
./Tools/Fold-Changelog.ps1
./Tools/Fold-Changelog.ps1 -Check
```

The shared behavior suite is [`Tools/Test-ChangelogWorkflow.ps1`](../../Tools/Test-ChangelogWorkflow.ps1).
Repository-specific adapters supply their existing test scratch root and register the suite with normal tooling
validation. The runner bounds its owned child and pipe waits and terminates its child tree on timeout before
fixture cleanup. This tooling-only workflow does not establish native runtime, platform, rendering or performance qualification.
