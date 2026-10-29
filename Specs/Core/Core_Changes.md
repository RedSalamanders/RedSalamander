# Changes workflow

Status: normative repository workflow contract
Last reviewed: 2026-10-29

DxUi, RedSalamander, RedXe, RedPrism, RedTilde and RedMonitor use the same fragment format, commands and behavioral
fixtures. Each repository owns its local changelog; changing one repository never updates a consumer's notes or
dependency pin.

Record each change as one UTF-8 Markdown file named `Changes/<yyyy-mm-dd>-<topic>.md`, using a real calendar date
and a unique topic in lowercase letters, digits and hyphens. It contains one nonempty top-level `- ` bullet;
indent all continuation lines and sub-bullets with spaces. Describe the behavior, relevant validation and material
remaining gates. Do not maintain a second manual list or hand-edit generated `CHANGELOG.md` entries.

**A change keeps its fragment; fragments are folded on the default branch after merge.** A feature change never
folds and never edits `CHANGELOG.md`, so parallel changes never conflict on the changelog. After merges, an
integration owner runs `Tools/Submit-ChangelogFold.ps1`, which folds the default branch's fragments in a pull
request of its own; merge it before the next fold. Existing release notes remain in place; adopting this workflow
does not migrate their history.

[`Tools/Test-Changelog.ps1`](../../Tools/Test-Changelog.ps1) is the read-only validator. It rejects invalid
names/dates, empty or multiple bullets, unindented continuation lines, unexpected entries under `Changes/`, reparse
paths inside the repository, invalid UTF-8 (only a UTF-8 byte order mark is accepted; a UTF-16 or UTF-32 one is not
honored), missing `Changes/README.md`, and a missing or duplicate `## Unreleased` heading in `CHANGELOG.md`. With
`-Base <ref>` it also compares the checkout with its merge base with that ref: fragments may be added, edited or
renamed, and `CHANGELOG.md` must equal the exact fold `Fold-Changelog.ps1` would write of the fragments that the merge
base holds and the checkout no longer has. Hand edits, entries folded from fragments the same change adds, edited,
removed, reordered or duplicated entries, any other change outside the Unreleased entries, and fragments removed
without their entry are refused; line endings never count. PrePush entrypoints run it with `-Base origin/main` and CI with `-Base HEAD^1`
from a two-commit checkout; repositories without PrePush run the same command before handoff. Existing
validation/tooling entrypoints invoke it or its shared module. Explanation-only runs never mutate files.

[`Tools/Fold-Changelog.ps1`](../../Tools/Fold-Changelog.ps1) validates all input before writing. It inserts entries
under Unreleased in descending ordinal filename order (newest date first, topic breaking ties), preserves the
changelog's UTF-8 BOM and line endings, flushes a staged changelog to disk and atomically publishes it, then removes
only folded files. A removal that fails after publication names the fragment; rerunning finishes the cleanup. A
checkout-local exclusive lease prevents overlapping folds; repositories ignore its lease and staged files. Retry after
interrupted cleanup does not duplicate an identical entry already in Unreleased. An identical entry in a released
section does not suppress a new entry. `-WhatIf` previews without writing; `-Check` is read-only and rejects even one
remaining fragment, which makes it the check of a finished fold. Validation and folding never stage, commit, push,
fetch, build native code or take desktop focus. Reparse checks cover the repository root and everything below it,
not the folders above it. An explicit relative `-Root` resolves against the caller's PowerShell location, including
after `Set-Location`, independently of the process working directory. Validation, preview and folding all target
that same checkout.

[`Tools/Submit-ChangelogFold.ps1`](../../Tools/Submit-ChangelogFold.ps1) fetches the default branch, folds it in a
temporary worktree outside the caller's checkout, checks the result with `Test-Changelog.ps1 -Base`, commits with the
repository's own identity, signing and hooks, pushes `changelog/fold-<commit>` and opens its pull request with `gh`
(`-PushOnly` leaves the pull request to the caller). The temporary worktree goes under `-WorktreeParent` (the user's
temporary folder by default; test harnesses pass a folder in their sandbox). It removes only its own temporary
worktree, never pruning other registrations, and changes nothing in the caller's checkout. It pushes nothing when
there is nothing to fold, reuses a branch already pushed for the same commit, and reports an open fold pull request,
looking through every open pull request, instead of opening a second, conflicting one. `-WhatIf` folds and commits
only in the temporary worktree and never pushes or opens a pull request, also for a fold branch pushed earlier. GitHub
Actions never fold: they cannot open pull requests in these repositories, and private repositories' hosted jobs are
not always available.

```powershell
./Tools/Test-Changelog.ps1
./Tools/Test-Changelog.ps1 -Base origin/main
./Tools/Submit-ChangelogFold.ps1 -WhatIf
./Tools/Submit-ChangelogFold.ps1
```

The shared behavior suite is [`Tools/Test-ChangelogWorkflow.ps1`](../../Tools/Test-ChangelogWorkflow.ps1). It uses
local Git fixture repositories with a fixed identity and none of the person's hooks or signing, a local bare remote
and a stand-in `gh` for the submission, and keeps every temporary worktree inside its scratch root. Controls that need
Windows (junctions, delete sharing) report SKIP on other hosts. Repository-specific adapters supply their existing
test scratch root and register the suite with normal tooling validation. The runner bounds its owned child and pipe waits and terminates its child tree
on timeout before fixture cleanup. This tooling-only workflow does not establish native runtime, platform, rendering
or performance qualification.
