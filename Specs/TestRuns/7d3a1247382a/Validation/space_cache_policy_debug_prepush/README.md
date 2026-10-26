# Scoped validation evidence

Run `20261009T110605Z-13812-2d05a89c3071489ba13d840530fd532c`; mode `Fresh`, requested gate `Full`, scope `filtered`. Repository verdict: **FAILED**; change verdict: **FAILED**. Cases: 1185 passed, 6 failed, 28 skipped of 1219.

This is the emitted scoped aggregate, not a claim of Full repository qualification. Preserve the exact emitted records, including every executed/reused/provisional/promoted/incomplete entry and every case status/reason.

- Plan digest: `4084a99273b238a7d028769d47563005b8f77547bdf1e6f2b2d8a5b07891fb9b`; coverage plan: `1c7454e573fbba1e513bc5af7a9e0fceefbabd98f1905b1d12d214470a5f68c8`.
- Decision digest: `89a2c3853cc749bf47a9f0cc9f1e4b89a15ac89f6a7ab297d856587006f352e5`; fingerprint record digest: `e3a9f5e2b32ee6ecf7a49f0cde904ac5c9d884c22675a8b183c301dc557e1264`.
- Initial workspace snapshot: `1034099f57e1ecfce65f0074dfb89a5b9955ca067d9abee59070aaa21e002d28`; final snapshot: `1034099f57e1ecfce65f0074dfb89a5b9955ca067d9abee59070aaa21e002d28`.
- Build receipt: `2bf8224f970817fc029a478e7bcb4a838e959c32050bb7afbaea2ebdd5468647`; platform/configuration: `x64/Debug`.
- Build source snapshot identity: `f43deb78049356ccb5e5fff99ba0f9f90b355b24d62a985b7df124c77ce2aea9`; v2 validation workspace snapshot: `1034099f57e1ecfce65f0074dfb89a5b9955ca067d9abee59070aaa21e002d28` (different record projections).
- Run-state terminal status: `failed`; mutation state: `stable`.
- Final artifact epoch: `1`; fingerprint record epoch: `1`.
- Executed/reused/provisional/promoted/invalidated/incomplete: 15/0/0/13/0/2.

- Detailed case history: 1207 rows; exact status counts {'passed': 1174, 'failed': 5, 'skipped': 28}, matching all 1207 v1 detailed-case rows. The v2 aggregate reports 1219 outcomes because it also counts 12 standalone suite/process outcomes that have no detailed case rows.
- Provenance: for this run, `previous-validation-plan.json`, `previous-validation-plan-decision.json`, `previous-entry-fingerprints.json`, and `workspace-initial.json` were found under `artifacts/selftest/previous_run`; their plan/coverage/decision/snapshot/fingerprint epoch and build-receipt references were checked against this run's v2 aggregate and terminal state before copying. `workspace-final.json` and the aggregate/case-history/v1 records are from `last_run`.

- Missing per-suite output logs referenced by v1 (not substituted with stale `previous_run` files): `D:\RedSalamander.Perf\runs\20261009T110605Z-13812-2d05a89c3071489ba13d840530fd532c\artifacts\selftest\last_run\CrashHandlingTests.output.log`; `D:\RedSalamander.Perf\runs\20261009T110605Z-13812-2d05a89c3071489ba13d840530fd532c\artifacts\selftest\last_run\LocalizationTests.output.log`; `D:\RedSalamander.Perf\runs\20261009T110605Z-13812-2d05a89c3071489ba13d840530fd532c\artifacts\selftest\last_run\MonitorTest.output.log`; `D:\RedSalamander.Perf\runs\20261009T110605Z-13812-2d05a89c3071489ba13d840530fd532c\artifacts\selftest\last_run\PluginContractTests.output.log`; `D:\RedSalamander.Perf\runs\20261009T110605Z-13812-2d05a89c3071489ba13d840530fd532c\artifacts\selftest\last_run\ProductUiTests.output.log`; `D:\RedSalamander.Perf\runs\20261009T110605Z-13812-2d05a89c3071489ba13d840530fd532c\artifacts\selftest\last_run\RedConfigureTests.output.log`; `D:\RedSalamander.Perf\runs\20261009T110605Z-13812-2d05a89c3071489ba13d840530fd532c\artifacts\selftest\last_run\SettingsSchemaTests.output.log`; `D:\RedSalamander.Perf\runs\20261009T110605Z-13812-2d05a89c3071489ba13d840530fd532c\artifacts\selftest\last_run\ViewerPETests.Noninteractive.output.log`; `D:\RedSalamander.Perf\runs\20261009T110605Z-13812-2d05a89c3071489ba13d840530fd532c\artifacts\selftest\last_run\ViewerSqliteTests.Noninteractive.output.log`. The raw outer launcher log is preserved and contains the emitted combined runner output.

## Entry dispositions

| Entry | Decision | Run-state | Aggregate execution/promotion | Reasons |
|---|---|---|---|---|
| `tooling.pester-build-toolchain` | EXECUTE | promoted | executed/promoted | SOURCE_INPUT_CHANGED |
| `selftest.compare-directories` | EXECUTE | promoted | executed/promoted | SOURCE_INPUT_CHANGED |
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

Payload files: 31; exact copied payload bytes: 2889107. `manifest.txt` records each source path, byte length, and SHA-256. All payloads were copied as raw bytes. `.gitattributes` disables text normalization for this archive.
