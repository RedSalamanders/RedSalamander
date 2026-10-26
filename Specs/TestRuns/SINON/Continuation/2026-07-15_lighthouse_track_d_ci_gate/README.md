# Lighthouse Track D — PR verification gate

Closeout evidence for Lighthouse `LH-9`, captured on 2026-07-15 from
`codex/lighthouse-ci-gate` based on commit `35119a7c9` plus the Track D working-tree change.

## Contract and baseline

- PluginContractTests, SettingsSchemaTests, and CrashHandlingTests each passed twice before the runner-plan
  edit: 6/6 direct Debug x64 executable invocations.
- `RunAllTestsPlan.Tests.ps1` passed 33/33 and guards both the three Suite CI entries and the explicit reusable
  Debug ARM64 build-only workflow lane.
- The reusable workflow selects a triplet-specific vcpkg cache key, so `arm64-windows` does not collide with
  the x64 cache in GitHub Actions.

## Build evidence

- Debug ARM64 build: `msbuild-20260715_113113_607.log`, 10m33s, 0 errors, 80 existing C4324
  alignment-padding warning lines.
- Debug x64 CI build: `msbuild-20260715_114319_235.log`, 0 warnings and 0 errors.
- The first attempted x64 CI build after the local ARM64 build was invalid environment evidence: the shared
  local x64 vcpkg merge had lost 2,902 files and compilation could not find WIL. `vcpkg-install.ps1 -Platform
  x64` restored and verified the triplet before the valid CI runs. This local shared-install interaction does
  not apply to the workflow's triplet-separated jobs and caches.

## CI compositions

| Run | Build | Result | New Track D adapters |
|-----|-------|--------|----------------------|
| `20260715T094318Z-60932-adb73e0932ee43dea10caddd1891578e` | Debug x64, fresh | 1,168 passed, 1 failed, 53 skipped; ViewerPETests `REGRESSION` | all 3 passed |
| `20260715T102731Z-69604-cc848739f552462a82eaf5775d54c352` | `-SkipBuild` | 1,169 passed, 0 failed, 53 skipped; all 33 suites passed | all 3 passed |

The first run's ViewerPETests failure reproduced on its runner retry in
`TestViewerImgRawLatestWinsExactReaderAndCloseSafety`; the second full CI composition passed ViewerPETests.
This is the independently reproduced Farsight latest-wins contract regression already observed during
Lighthouse Track B. It remains blocking external evidence and was not quarantined or made non-blocking.

`ci_pass_1/` preserves the aggregate plus both ViewerPETests logs. `ci_pass_2/` preserves the green aggregate.
`builds/` preserves the x64 and ARM64 MSBuild logs.
