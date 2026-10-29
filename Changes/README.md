# Changelog fragments

Record each change as `Changes/<yyyy-mm-dd>-<topic>.md`: a real date, a unique lowercase topic using letters,
digits and hyphens, and one nonempty Markdown `- ` bullet. Indent every continuation line and sub-bullet
with spaces. Describe behavior, relevant validation and material remaining gates. Use UTF-8.

A change keeps its fragment: do not fold it and do not edit CHANGELOG.md, so parallel branches add distinct
files and never conflict on the changelog. After changes merge, fold the default branch's fragments in a pull
request of its own. Do not hand-edit generated CHANGELOG.md entries or maintain a second manual list.

```powershell
./Tools/Test-Changelog.ps1                     # while editing: fragment names and content
./Tools/Test-Changelog.ps1 -Base origin/main   # before a push (PrePush runs it): CHANGELOG.md changes only by exact folds
./Tools/Submit-ChangelogFold.ps1 -WhatIf       # after merges: preview the fold of the default branch
./Tools/Submit-ChangelogFold.ps1               # fold it on changelog/fold-<commit> and open that pull request
```

The submission folds the default branch in a temporary worktree, never in your checkout, and pushes nothing when
there is nothing to fold or a fold pull request is already open; merge that one first. Folding puts entries under
CHANGELOG.md's single `## Unreleased` heading in descending filename order and removes only folded files. It
preserves UTF-8 BOM and line endings; rerunning after an interrupted cleanup finishes it without duplicating
published entries. With `-Base`, validation refuses hand edits, entries folded from fragments the same change adds,
edited entries and fragments removed without their entry. Validation and `Tools/Fold-Changelog.ps1` never stage,
commit, push or fetch; only the submission does, and no command takes desktop focus.

See the [normative workflow](../Specs/Core/Core_Changes.md). Existing release notes remain in place.
