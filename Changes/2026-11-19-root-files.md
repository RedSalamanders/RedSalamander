- Converged the root files with every RedSalamanders repository (convergence CV-6-03, owner decision D18 and the
  B-12 default): `LICENSE.txt` is now only the common MIT text, and its third-party list moves verbatim to
  `THIRD-PARTY-NOTICES.md`, which also names the shipped components the list predates. The portable ZIP builder
  copies the notices file beside `LICENSE.txt` and fails without it. The common `CONTRIBUTING.md`, pull request
  template and `.vsconfig` are added; the template's RedSalamander item names move to `CONTRIBUTING.md`. Validated
  with `Tools/Get-SpecInventory.ps1 -FailOnFindings -NoReports` and `Tools/Test-Changelog.ps1`; native packaging runs
  in main qualification.
