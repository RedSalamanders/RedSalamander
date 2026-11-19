- Converged part of the specifications with every RedSalamanders repository (convergence CV-6-03): the common
  `Build_UnifiedCI.md` skeleton, dated names for the two undated active plans
  (`FileSystem_GoogleDrivePluginPlan_2026-03-08.md`, `Win32_Inventory_2026-05-18.md`), the four common plan statuses
  (the unused `DRAFT` label is retired), the common CODEOWNERS file, and a RedMonitor consumer contract that keeps only
  RedSalamander's adapter facts and links to RedMonitor's. Rules moved unchanged; no behavior, build input or test
  changed. Validated with `Tools/Get-SpecInventory.ps1 -FailOnFindings -NoReports`, the specification Pester
  contracts and `Tools/Test-Changelog.ps1`.
