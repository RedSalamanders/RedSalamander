- Skill packages no longer carry Codex UI metadata (`RedHome` convergence step CV-2-11, register A-11 decided
  "none"): the three `agents/openai.yaml` files of `red-salamander-crash-forensics`, `build-preferences-pages` and
  `tooling-governance` are removed. Nothing in the repository reads them; historical run records keep their copies.
  The 11 project skills move from `.github/skills/` to `.agents/skills/`, the skill folder of every RedSalamanders
  repository and the one Codex, Antigravity and Copilot read (CV-2-09, register A-09), and drop their
  `metadata.author` and `metadata.version` keys so front matter is `name` and `description` only (A-10). `AGENTS.md`,
  `Specs/Testing/Testing_AgentGuidance.md`, the `Specs/NormativeConsistency.json` anchors, two active plans and the
  default guidance roots of `Tools/Tests/DocumentationDriftContracts.Tests.ps1` follow the move.
  `red-salamander-crash-forensics`, `cpp-build` and `tooling-governance` now say when to use them ("Use when ...",
  CV-2-13), so every skill description says what the skill does and then when to use it.
