# Changes workflow

Status: normative repository workflow contract
Last reviewed: 2026-11-19

The fragment format, the fold and the commands are owned once by the shared CI engine's
[project commands contract](https://github.com/RedSalamanders/.github/blob/f965c7945aff98dbab352623a0cfbe333e7f282f/Specs/ProjectCommands.md), at the
commit `.ci/project.json` pins. This repository's `Tools/Test-Changelog.ps1`, `Tools/Fold-Changelog.ps1`,
`Tools/Submit-ChangelogFold.ps1`, `Tools/Release-Changelog.ps1` and `Tools/Test-ChangelogWorkflow.ps1` are that
engine's shims: the same names and parameters, running the implementation in `.ci-engine`. They refuse with one
message that names the fix when `.ci-engine` is missing, at another commit or changed. Each repository owns its local
changelog; changing one repository never updates a consumer's notes or dependency pin.

Record each change as one UTF-8 Markdown file named `Changes/<yyyy-mm-dd>-<topic>.md`, using a real calendar date
and a unique topic in lowercase letters, digits and hyphens. It contains one nonempty top-level `- ` bullet;
indent all continuation lines and sub-bullets with spaces. Describe the behavior, relevant validation and material
remaining gates. Do not maintain a second manual list or hand-edit generated `CHANGELOG.md` entries.

**A change keeps its fragment; fragments are folded on the default branch after merge.** A feature change never
folds and never edits `CHANGELOG.md`, so parallel changes never conflict on the changelog. After merges, an
integration owner runs `Tools/Submit-ChangelogFold.ps1`, which folds the default branch's fragments in a pull
request of its own; merge it before the next fold. Existing release notes remain in place.

In this repository, `test.ps1` runs `Tools/Test-Changelog.ps1` before execution. PrePush passes
`-Base origin/main` or its explicit `-BaseRef`. `Test-Changes.ps1` forwards existing arguments to `test.ps1`.
`-Explain` only produces the test plan. `Tools/Run-AllTests.ps1` runs the validator in its tooling lane.
The Pester suite `Tools/Tests/Changelog.Tests.ps1` runs `Tools/Test-ChangelogWorkflow.ps1` in the marked test sandbox. The
compiler-free PR checks pass the immutable PR base SHA to `Tools/Invoke-ProjectCi.ps1 -Phase Static`, which runs it
with `-Base <base SHA>`. The behavior suite also proves that these shims equal the pinned engine's templates and that
the PrePush runner checks the changelog first. This tooling-only workflow does not establish native runtime,
platform, rendering or performance qualification.

```powershell
./Tools/Test-Changelog.ps1
./Tools/Test-Changelog.ps1 -Base origin/main
./Tools/Submit-ChangelogFold.ps1 -WhatIf
./Tools/Submit-ChangelogFold.ps1
```
