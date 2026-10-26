# Scoped validation evidence

Run `20261009T070512Z-41240-3056aebd18704c8aaea868054b1b9ba0`; mode `Fresh`, requested gate `Full`, scope `filtered`. Repository verdict: **FAILED**; change verdict: **FAILED**. Cases: 1182 passed, 9 failed, 28 skipped of 1219.

This is the emitted scoped aggregate, not a claim of Full repository qualification. Preserve the exact emitted records, including every executed/reused/provisional/promoted/incomplete entry and every case status/reason.

- Plan digest: `4084a99273b238a7d028769d47563005b8f77547bdf1e6f2b2d8a5b07891fb9b`; coverage plan: `1c7454e573fbba1e513bc5af7a9e0fceefbabd98f1905b1d12d214470a5f68c8`.
- Decision digest: `c07e930fb7ec91d6cb2da63dabb25f13c409f0a5c86831ffb00db7db8a987857`; fingerprint record digest: `b27b4e2f87caf863d5bebc83815e1e80863bc11d73d38b8a925ab1bfb44bbe4d`.
- Initial workspace snapshot: `d0558ea74aa717ae96028bf3ad4445d9b1ce3398171f50bd768b9cbcb96d59d7`; final snapshot: `d0558ea74aa717ae96028bf3ad4445d9b1ce3398171f50bd768b9cbcb96d59d7`.
- Build receipt: `789bd717926d71ba8c6f571613fc6ad0d0a8fda8a35174c811d06c153587db34`; platform/configuration: `x64/Debug`.
- Build source snapshot identity: `b890ae0ddbf01b41819c00e27abd4c803a32a179860ac1d57da68da7266a80aa`; v2 validation workspace snapshot: `d0558ea74aa717ae96028bf3ad4445d9b1ce3398171f50bd768b9cbcb96d59d7` (different record projections).
- Run-state terminal status: `failed`; mutation state: `stable`.
- Final artifact epoch: `1`; fingerprint record epoch: `1`.
- Executed/reused/provisional/promoted/invalidated/incomplete: 15/0/0/12/0/3.

- Detailed case history: 1207 rows; exact status counts {'passed': 1171, 'failed': 8, 'skipped': 28}, matching all 1207 v1 detailed-case rows. The v2 aggregate reports 1219 outcomes because it also counts 12 standalone suite/process outcomes that have no detailed case rows.
- Provenance: `previous-validation-plan.json`, `previous-validation-plan-decision.json`, `previous-entry-fingerprints.json`, and `workspace-initial.json` were found for this run under `artifacts/selftest/previous_run`; their plan/coverage/decision/snapshot/fingerprint epoch and build-receipt references were checked against this run's v2 aggregate and terminal state before copying. `workspace-final.json` and the aggregate/case-history/v1 records are from `last_run`.

- Missing per-suite output logs referenced by v1 (not substituted with stale `previous_run` files): `D:\RedSalamander.Perf\runs\20261009T070512Z-41240-3056aebd18704c8aaea868054b1b9ba0\artifacts\selftest\last_run\CrashHandlingTests.output.log`; `D:\RedSalamander.Perf\runs\20261009T070512Z-41240-3056aebd18704c8aaea868054b1b9ba0\artifacts\selftest\last_run\LocalizationTests.output.log`; `D:\RedSalamander.Perf\runs\20261009T070512Z-41240-3056aebd18704c8aaea868054b1b9ba0\artifacts\selftest\last_run\MonitorTest.output.log`; `D:\RedSalamander.Perf\runs\20261009T070512Z-41240-3056aebd18704c8aaea868054b1b9ba0\artifacts\selftest\last_run\PluginContractTests.output.log`; `D:\RedSalamander.Perf\runs\20261009T070512Z-41240-3056aebd18704c8aaea868054b1b9ba0\artifacts\selftest\last_run\ProductUiTests.output.log`; `D:\RedSalamander.Perf\runs\20261009T070512Z-41240-3056aebd18704c8aaea868054b1b9ba0\artifacts\selftest\last_run\RedConfigureTests.output.log`; `D:\RedSalamander.Perf\runs\20261009T070512Z-41240-3056aebd18704c8aaea868054b1b9ba0\artifacts\selftest\last_run\SettingsSchemaTests.output.log`; `D:\RedSalamander.Perf\runs\20261009T070512Z-41240-3056aebd18704c8aaea868054b1b9ba0\artifacts\selftest\last_run\ViewerPETests.Noninteractive.output.log`; `D:\RedSalamander.Perf\runs\20261009T070512Z-41240-3056aebd18704c8aaea868054b1b9ba0\artifacts\selftest\last_run\ViewerSqliteTests.Noninteractive.output.log`. The raw outer launcher log is preserved and contains the emitted combined runner output.

## Entry dispositions

| Entry | Decision | Run-state | Aggregate execution/promotion | Reasons |
|---|---|---|---|---|
| `tooling.pester-build-toolchain` | EXECUTE | promoted | executed/promoted | SOURCE_INPUT_CHANGED |
| `selftest.compare-directories` | EXECUTE | failed | executed/incomplete | SOURCE_INPUT_CHANGED |
| `standalone.product-ui` | EXECUTE | promoted | executed/promoted | SOURCE_INPUT_CHANGED |
| `viewer-pe.noninteractive` | EXECUTE | promoted | executed/promoted | SOURCE_INPUT_CHANGED |
| `viewer-sqlite.noninteractive` | EXECUTE | promoted | executed/promoted | SOURCE_INPUT_CHANGED |
| `standalone.filesystem-curl` | EXECUTE | promoted | executed/promoted | SOURCE_INPUT_CHANGED |
| `standalone.monitor` | EXECUTE | promoted | executed/promoted | SOURCE_INPUT_CHANGED |
| `standalone.localization` | EXECUTE | promoted | executed/promoted | SOURCE_INPUT_CHANGED |
| `standalone.red-configure` | EXECUTE | promoted | executed/promoted | SOURCE_INPUT_CHANGED |
| `standalone.plugin-contract` | EXECUTE | failed | executed/incomplete | SOURCE_INPUT_CHANGED |
| `standalone.settings-schema` | EXECUTE | promoted | executed/promoted | SOURCE_INPUT_CHANGED |
| `standalone.crash-handling` | EXECUTE | promoted | executed/promoted | SOURCE_INPUT_CHANGED |
| `perf.performance-tests2` | EXECUTE | promoted | executed/promoted | SOURCE_INPUT_CHANGED |
| `tooling.pester` | EXECUTE | failed | executed/incomplete | SOURCE_INPUT_CHANGED |
| `tooling.vcpkg-merge-synthetic` | EXECUTE | promoted | executed/promoted | SOURCE_INPUT_CHANGED |

## Pending interactive and delegated CI

- `PENDING_INTERACTIVE (5): Commands, FileOperations, ViewerPETests.Interactive, ViewerSqliteTests.Interactive, RedSalamanderMonitorEtwLatency`

No `DEFERRED_CI` line was found in the preserved launcher log.

Broad `viewer-pe.noninteractive` executed. Its v1-referenced per-suite stdout file was absent from this run's `last_run`; the combined outer launcher stream and all eight lifecycle PNG/provenance pairs are preserved. Whole-frame review is pending unless separately recorded by the reviewer.

Payload files: 31; exact copied payload bytes: 3438749. `manifest.txt` records each source path, byte length, and SHA-256. All payloads were copied as raw bytes. `.gitattributes` disables text normalization for this archive.

## Failure details

The eight detailed failed rows are preserved in the v1 aggregate, case-history file, and combined launcher log. The ninth v2 failure is the standalone `PluginContractTests` process outcome, which has no detailed case-history row.

- `CompareDirectories/local_search_callback_contract` failed after 8 ms: `Search abort-match case expected ERROR_CANCELLED. hr=0x00000000`.
- `ToolsPesterTests` had seven failures and one skip. Five link/reparse fixture cases failed in `New-RSTestSandboxLink` with `A required privilege is not held by the client` / `Administrator privilege required for this operation`: `rejects a final-file reparse destination without changing its external target`; `rejects a real directory symbolic-link ancestor`; `rejects a real file symbolic link and a dangling symbolic-link component`; `refuses a test link that would live or point outside its marked sandbox, before creating it`; and `rejects reparse inputs and unsafe impact-base arguments without fetching`.
- The other two Pester failures were source contracts: `centralizes callback-return module-pin transfer` rejected the direct `FreeLibraryWhenCallbackReturns(instance, context->modulePin.release())` use, expecting the shared `TransferModulePinToCallbackReturn` helper; `finds no production GetCursorPos() usage outside SelfTest/Tests sources` found `Plugins/ViewerSpace/ViewerSpace.cpp:5662`.
- `PluginContractTests` failed overall. Its emitted log reports `FileSystem.dll` failed one debug selftest because the relative file-symlink fixture could not be created (the log states `SeCreateSymbolicLinkPrivilege` is not held); `FileSystemS3.dll` failed three checks: recursive Delete did not list the fixture prefix, streaming cancellation did not return while the body was arriving, and cancellation exceeded the callback bound with a 15000 ms return. The archive preserves the outer runner stream because the v1-referenced per-suite output file is absent from this run's `last_run`; no prior-run log was substituted. These remain failures, not accepted environmental skips.
