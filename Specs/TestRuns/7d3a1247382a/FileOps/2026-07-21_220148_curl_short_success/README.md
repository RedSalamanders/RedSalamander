# Curl short-success staged-upload proof

Scope: FG-P0-3 proof depth, x64, 2026-07-21, machine hash `7d3a1247382a`.

The test-only fake FTP endpoints bind ephemeral IPv4 loopback ports and drive the public cross-connection FTP
`MoveItem(...)` path. Both endpoints consume the complete Curl upload and return `226`; the fault endpoint retains
only a strict prefix. The production upload, targeted staged `SIZE`, cleanup, promotion, overwrite rollback, and source
delete paths are therefore exercised without an external service or production bypass.

## Results

- Debug `FileSystemCurl` build: `msbuild-20260721_225539_897.log`, 0 warnings, 0 errors.
- Debug `PluginContractTests`: Curl provider selftests 102/102; complete harness passed.
- Test-enabled Release full recovery rebuild: `msbuild-20260721_234225_285.log`, 0 errors. Its 10 `/Wall` warnings
  are pre-existing `ViewerPETests` unused-symbol warnings; none is in FileSystemCurl. Test-enabled Release
  `PluginContractTests` passed with Curl provider selftests 102/102.
- Production tests-disabled Release `FileSystemCurl` rebuild: `msbuild-20260721_235846_645.log`, 0 warnings, 0
  errors. `dumpbin /exports` found zero `RedSalamanderCurlDebugSelfTests` exports, and the production Release
  `PluginContractTests` run passed with the Curl selftest step explicitly skipped.
- Standalone `FileSystemCurlTests` passed in Debug and Release. Relevant Pester source contracts passed 149/149.

## Exact transport evidence

- Short-success case: 4,099 upload bytes consumed, 2,049 retained, 15 destination commands, 20,000 us, exactly one
  staged `STOR`, exactly one staged `SIZE`, staged sibling deleted, existing final bytes unchanged, source bytes
  unchanged, no source `DELE`, result `0x8007012B` (`ERROR_PARTIAL_COPY`).
- Complete-upload control: 4,099 bytes consumed and retained, 27 destination commands, 30,000 us, one staged
  `STOR` and `SIZE`, exact promotion over the existing destination, rollback sibling cleanup, exactly one source
  `DELE`, result `S_OK`.
- Both cases are code-gated to at most 15 seconds, source command count at most 64, and destination command count at
  most 96 for the short case / 128 for the control.

## Mutation sensitivity

The staged-size comparison was temporarily changed from `stagedProbeSize != fileSize` to
`stagedProbeSize > fileSize`, which bypasses a retained-prefix mismatch. The Debug plugin rebuilt warning-clean, but
the provider result became 97 passed / 5 failed with `E_FAIL`, and `PluginContractTests` exited 1. Restoring `!=`
returned the same harness to 102/102 and exit 0. The production guard is restored in the archived source state.

The durable behavior contract is in `Specs/FileSystem/FileSystem_FileOperations.md`; the test-build and deterministic
fault-injection contract is in `Specs/Testing/Testing_SelfTests.md`.
