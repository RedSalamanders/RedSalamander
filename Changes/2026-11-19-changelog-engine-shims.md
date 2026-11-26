- Run the changelog commands from the pinned shared CI engine (convergence CV-5-02). `Tools/Test-Changelog.ps1`,
  `Fold-Changelog.ps1`, `Submit-ChangelogFold.ps1`, `Test-ChangelogWorkflow.ps1` and the new `Release-Changelog.ps1`
  are now copies of the engine's shims (`Templates/Project/Tools` at `f965c79`): the same names and parameters, running
  the engine implementation in `.ci-engine`, and refusing with one message when that checkout is missing, at another
  commit or changed. `Tools/Modules/Documentation/Changelog.psm1` is removed; the engine implementation keeps this
  repository's refusal of two pending fragments with the same text and adds DxUi's checks (a Git error while reading
  the base fails `-Base`; a reused fold branch must be the exact fold; the exact release transition is accepted). The
  scoped-testing fixture gets its own manifest and `.ci-engine` clone; the tool inventory, the normative-consistency
  anchors and `Specs/Core/Core_Changes.md` follow. Tooling-only; native runtime coverage is unchanged.
