# Microsoft Drive ambiguous mutation reconciliation evidence

Scope: FG-P2-1 and FG-P2-2, x64, 2026-07-21.

- `plugin-contract-debug.txt`: complete Debug `PluginContractTests`; Microsoft Drive provider selftests pass
  165/165 and the executable exits successfully.
- `plugin-contract-test-enabled-release.txt`: complete test-enabled Release `PluginContractTests`; Microsoft Drive
  provider selftests pass 165/165 and the executable exits successfully.
- Focused provider build logs: `msbuild-20260721_142041_783.log` (Debug) and
  `msbuild-20260721_142147_870.log` (test-enabled Release), both zero warnings/errors.
- Production tests-disabled provider build: `msbuild-20260721_142254_835.log`, zero warnings/errors; final production
  Release rebuild `msbuild-20260721_193332_845.log`; `dumpbin /exports` reports zero
  `RedSalamanderMicrosoftDriveDebugSelfTests` exports.
- Existing source contracts pass 149/149.

The provider fake-Graph ledger proves the following request shapes:

- ordinary successful PATCH: one PATCH, zero by-ID reconciliation GETs;
- ordinary successful DELETE: one DELETE, zero by-ID reconciliation GETs;
- failed PATCH transport: at most one by-ID GET, accepted only for exact stable ID, name, and requested parent;
- failed DELETE transport: at most one by-ID GET, accepted only when the stable item ID is absent;
- first DELETE commits but returns `503`, retry receives `404/itemNotFound`: two DELETE attempts, zero by-ID GETs,
  operation succeeds;
- definitive HTTP conflict remains mapped as failure and adds no reconciliation GET.

The durable behavior contract is in `Specs/FileSystem/FileSystem_MicrosoftDrive.md`.
