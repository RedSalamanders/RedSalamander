- Consolidated the shared `improve`, `compiler-warnings`, `cpp-modern-style`, `error-handling`, `localization`,
  `wil-raii`, and `yyjson` agent skills into the organization `RedHome` repository, retaining the twelve
  RedSalamander-specific skills and validating with `Tools/Get-SpecInventory.ps1 -FailOnFindings`,
  `DocumentationDriftContracts.Tests.ps1`, and `Tools/Test-Changelog.ps1` (documentation/tooling only; native
  qualification unchanged).
  Shared-skill discovery names `../RedHome/AGENTS.md`, the package directory, and the clone URL while preserving
  local product authority; the SettingsStore overview now reflects its JSON/JSON5 persistence contract.
