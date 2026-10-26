# Riptide R2-2 Staging Name Entropy

Scope: prove S3 hidden backup keys and curl remote rollback/upload sibling paths no longer use predictable PID, tick-count, or sequence-derived names.

## RED

- Build: `.\build.ps1 -ProjectName FileSystemS3 -Configuration Debug`
  - Log: `RED-build-FileSystemS3-msbuild-20260615_191703_293.log`
  - Result: built with the new debug selftest export before the production fix. The temporary red helper produced warning noise, fixed in GREEN.
- Build: `.\build.ps1 -ProjectName FileSystemCurl -Configuration Debug`
  - Log: `RED-build-FileSystemCurl-msbuild-20260615_191802_724.log`
  - Result: built with the behavior-preserving old curl sibling-name helper extraction before the production fix. The temporary red helper produced warning noise, fixed in GREEN.
- Build: `.\build.ps1 -ProjectName PluginContractTests -Configuration Debug`
  - Log: `RED-build-PluginContractTests-msbuild-20260615_191827_587.log`
  - Result: 0 warnings, 0 errors.
- Contract run: `.\.build\x64\Debug\PluginContractTests.exe`
  - Result: failed for the intended reason in Step5 only.
  - S3 debug selftests reported `passed=44, failed=24`, catching the old `.rs-bak-<pid>-<tick>-<counter>-<leaf>` shape.
  - Curl debug selftests reported `passed=52, failed=8`, catching the old 16-hex tick/sequence sibling leaf instead of a 128-bit entropy token.

## GREEN

- Build: `.\build.ps1 -ProjectName FileSystemS3 -Configuration Debug`
  - Log: `GREEN-build-FileSystemS3-msbuild-20260615_191930_536.log`
  - Result: 0 warnings, 0 errors.
- Build: `.\build.ps1 -ProjectName FileSystemCurl -Configuration Debug`
  - Log: `GREEN-build-FileSystemCurl-msbuild-20260615_192022_988.log`
  - Result: 0 warnings, 0 errors.
- Build: `.\build.ps1 -ProjectName PluginContractTests -Configuration Debug`
  - Log: `GREEN-build-PluginContractTests-msbuild-20260615_192048_012.log`
  - Result: 0 warnings, 0 errors.
- Contract run: `.\.build\x64\Debug\PluginContractTests.exe`
  - Artifact: `GREEN-PluginContractTests.txt`
  - Result: passed.
  - Debug selftests: Microsoft Drive `passed=25, failed=0`; S3 `passed=76, failed=0`; curl `passed=68, failed=0`.

## Implementation Notes

- `BuildHiddenSiblingKey(...)` now returns `HRESULT` and fills an output key from 16 bytes produced by `BCryptGenRandom`; RNG failure aborts the staging operation and rolls back prior S3 transfer journal work instead of falling back to predictable names.
- S3 hidden backup keys keep the parent and original leaf but replace PID/tick/counter material with a 32-character hex entropy token.
- Curl `GenerateRemoteSiblingPath(...)` now generates a fresh random 32-character hex token for every collision attempt instead of deriving candidates from `GetTickCount64()` plus a static sequence.
- `PluginContractTests` now runs optional Debug-only selftest exports for Microsoft Drive, S3, and curl; missing exports remain allowed for configurations that do not expose provider-local debug tests.
