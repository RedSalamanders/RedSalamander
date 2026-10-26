# Changelog fragments

Record each change as `Changes/<yyyy-mm-dd>-<topic>.md`: a real date, a unique lowercase topic using letters,
digits and hyphens, and one nonempty Markdown `- ` bullet. Indent every continuation line and sub-bullet
with spaces. Describe behavior, relevant validation and material remaining gates. Use UTF-8.

Keep fragments while editing so parallel branches can add distinct files. Fold before PrePush and final handoff,
and immediately after every merge/rebase that introduces fragments, even for one fragment. Do not wait for a
release, a large batch or a separate cleanup PR. Include the generated changelog and fragment removals in the
current change. Do not hand-edit generated CHANGELOG.md entries or maintain a second manual list.

```powershell
./Tools/Test-Changelog.ps1
./Tools/Fold-Changelog.ps1 -WhatIf
./Tools/Fold-Changelog.ps1
./Tools/Fold-Changelog.ps1 -Check
```

Validation is read-only. Folding puts entries under CHANGELOG.md's single `## Unreleased` heading in descending
filename order and removes only folded files. It preserves UTF-8 BOM and line endings; retrying interrupted
cleanup does not duplicate published entries. `-Check` rejects even one pending fragment. The commands never
stage, commit, push, fetch, build native code or take desktop focus.

See the [normative workflow](../Specs/Core/Core_Changes.md). Existing release notes remain in place.
