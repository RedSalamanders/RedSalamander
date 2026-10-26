# I28 Release dependency recovery

Status: DONE — recovery closed 2026-10-17. The build, pin, and shared-runtime cleanup below are the completed record. Unfinished Full qualification belongs to [I29](../WIP/FileOperations_PostMergeQualification_2026-10-05.md) and [I31](../WIP/DxUi_HostMessagesToolsetAndRestore_2026-10-01.md). This file does not claim a passing Full.
Owner: codex/fix-release-dependencies
Scope: restore the retained DxUi pin, repair the AWS dependency pair, and preserve shared runtimes through the native cleanup path blocking the x64 Release build.
Authority: [Build_Toolchain.md](../../Build/Build_Toolchain.md).

- [x] Inspect the restore log and preserve the incomplete consumer-managed vcpkg clone.
- [x] Restore DxUi `19ca44d95a0cd493fc7fb1a57f4d4ad227bdefc1` without changing its API-revision-2 lock or the sibling checkout.
- [x] Confirm the updater's divergent-history guard and current main's API revision 3; adoption remains separate.
- [x] Diagnose AWS SDK `1.11.880` compiling against CRT `0.43.0`; pin its upstream CRT `0.43.4` and align tool/registry identities.
- [x] Validate every manifest floor and override against the selected registry; run the existing dependency-tooling tests.
- [x] Reproduce shared DLL deletion with the installed native C++ cleanup task and protect its first-party output ownership.
- [x] Confirm the ordinary production Release command after native cleanup protection.
- [x] Run Fresh Full verification with native cleanup protection included and preserve its outcome.
- [x] Complete the x64 Release solution build and record its log and resulting receipt.
- [x] Run the required Fresh Full gate and preserve its qualification failures.
- [x] Route remaining Full failures to the owning contracts and retain their evidence.
- [x] Hand the unfinished Full qualification to [I29](../WIP/FileOperations_PostMergeQualification_2026-10-05.md) and [I31](../WIP/DxUi_HostMessagesToolsetAndRestore_2026-10-01.md). The October 2 Full runs stay failed. This recovery record does not rerun them and does not copy their checklists.

No controls, rendering, public DxUi API, or gallery pixels are changed by this work.
The incomplete clone is retained under the same consumer output directory with an `.interrupted-20261002-173459` suffix.

All manifest constraints exist at `d53f25d56e7806475054f624439e93ee38e1e81d`.
The governed installer compiled and merged all 36 packages, including AWS CRT `0.43.4`
and SDK `1.11.880` in Debug and Release. The focused install-safety and specification
information-architecture suites passed 47 tests with no skips; the specification inventory
had no blocking findings.

The production x64 Release solution build passed with one existing C5245 warning and no errors.
Its retained MSBuild log is `.build/logs/msbuild-20261002_200754_617-pid14168-15902e80.log`.
The test-enabled Full build and its full-solution re-attestation build also passed with no warnings
or errors; the latter log is `.build/logs/msbuild-20261002_204638_784-pid113008-8312c306.log`,
with receipt `756923f27f673595c657f752990cd7aaabfb1142123ba81353225c1c5d9907fd`.
Build receipts are bound to their source snapshot; this post-run record does not claim that the
earlier receipt attests later documentation edits.

The first ordinary production command compiled and linked, but attestation found missing
profile-root `sqlite3.dll`. Its log `.build/logs/msbuild-20261002_214545_088-pid11580-24c25988.log`
shows SearchService staging the DLL before later cleanup removed it. The previous shared-output
fixture modeled common MSBuild cleanup only. A new fixture imports the installed `Microsoft.CppClean.targets`
and real `CppClean` task: before the protection, the original nine tests passed while the native
case deleted another project's DLL; after the protection all ten tests passed without skips,
with owned and private outputs still removed. The ordinary `./build.ps1 -Configuration Release`
then passed with one existing warning and no errors, publishing receipt
`b5013c4d9517edd220560a82bc571b46f86045814fa3b517a583a24b48cf8a6f`.
Fresh Full verification with this cleanup protection completed in run
`20261002T201932Z-22808-3afb030086704778808523189b62c786`, using the ordinary environment,
with exit 1: 18 entry promotions and two failed entries. It reported 1,768 passes, two failures,
and 497 skips, including 468 cases left unexecuted by fail-fast. All 829 tooling tests passed
(13 build-toolchain and 816 other tooling tests), with no skips. ViewerPE, SQLite, Curl,
monitor ETW latency, and PerformanceTests2 passed. Compare Directories passed 243 cases with
29 environment-dependent skips. The ViewerPE close case measured 120 ms against 500 ms.

The Full build and both artifact-writer rebuilds passed with zero warnings and errors.
The final re-attestation log is `.build/logs/msbuild-20261002_223551_847-pid22808-bdd136a9.log`,
with receipt `f9314fea164f5f52d56ea26b0f8a24fcd8cf5dec650dd6d0b8e44a80fe2a7fb0`.
The run's source snapshot is `60c12aa8d213cc4b0ce8be46c0a097792695cb8c943ef09bb9d345840f5c30c7`;
later recovery-record and inventory-description edits are not attested by that earlier receipt.

- `Commands`: 608 passes, then
  `cmd_compare_directories_non_file_plugin_path_form_selection_and_empty_state`
  failed: selection inversion left `leftSelected=1`. Fail-fast left 324 cases unexecuted.
  Route behavioral diagnosis to [Compare Directories](../../Core/Core_CompareDirectories.md).
- `FileOperations`: 73 passes, then `FileOps_ProviderCapabilityMatrix` again exceeded
  the R0e Release ceiling: 2,499,220 us against 2,000,000 us. All 4,096 capability parses
  and admission decisions completed, one uncontained route was rejected, and provider calls
  remained zero. Admission work accounted for 2,463,414 us. Fail-fast left 144 cases unexecuted.
  Route measurement and optimization to [File Operations](../../FileSystem/FileSystem_FileOperations.md)
  and [performance validation](../../Testing/Testing_PerformanceValidation.md); retain the ceiling
  and require a paired baseline before claiming non-regression.

Both failures remain unclassified. The earlier conflict-prompt generation failure was not
re-executed because Commands stopped earlier. The [I27 safety-remediation record](UI_FileOperationsSafetyRemediation_2026-10-01.md) owns that
issue's resolution; this I28 record retains the original failed Full receipt and requires
independent final attribution.
This build-recovery record is closed. Full qualification of the failures below is no longer this plan's job; I29 and I31 own that Full. The later I26 Release Full `20261001T115346Z-104056-8e2cfb6a0b8049218c622dd915a2697e` (2,211/0/46) is a different source and does not close this recovery.

Fresh Full run `20261002T183108Z-113008-58bf908ec1094daba7d1b0bb1922857c` completed with exit 1.
Evidence is retained under `C:/RedSalamander.Perf/evidence/runs/<run-id>/`.
The full tooling group passed 816 tests with no skips. SDK plugin contracts, Curl, interactive
viewer groups, monitor ETW latency, and PerformanceTests2 passed. These three failures
stay open for I29 and I31. This record does not change their acceptance limits or claim a paired performance baseline:

- `ViewerPETests.Noninteractive`: `TestViewerPELatestWinsAndCloseDoesNotWaitForBlockedRead`
  measured Close at 1,447 ms against the 500 ms limit.
- `Commands`: 652 passes, then `cmd_pane_fileops_conflict_rejects_stale_same_task_presentation`
  failed to publish the distinct second prompt generation; fail-fast left 280 cases unexecuted.
- `FileOperations`: 73 passes, then `FileOps_ProviderCapabilityMatrix` measured 4,096 complete
  capability/admission decisions and zero provider calls at 2,441,242 us against the 2,000,000 us
  Release ceiling; fail-fast left 144 cases unexecuted.

The first Full attempt exhausted Z: during compilation. Completed task-created vcpkg caches
were preserved under the existing authorized test sandbox on C:; remaining entries were renamed
under `.build/cache-preservation/vcpkg-buildtrees-partial-20261002-183107`.
The retry built successfully. No source pins, installed libraries, or acceptance thresholds were
changed for space recovery. ARM64/ASan qualification and DxUi API-revision-3 adoption remain separate.
