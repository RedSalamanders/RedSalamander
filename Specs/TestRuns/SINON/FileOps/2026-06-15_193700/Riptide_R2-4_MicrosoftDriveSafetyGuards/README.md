# Riptide R2-4 Microsoft Drive Safety Guards

Scope: prove Microsoft Drive move/merge rejects self-descendant destinations, skips malformed empty-name children, and bounds recursive directory merges.

## RED

- Build: `.\build.ps1 -ProjectName FileSystemMicrosoftDrive -Configuration Debug`
  - Log: `RED-build-FileSystemMicrosoftDrive-msbuild-20260615_193511_515.log`
  - Result: 0 warnings, 0 errors. This build included the new debug selftests but not the production fix.
- Contract run: `.\.build\x64\Debug\PluginContractTests.exe`
  - Result: failed for the intended reason in Step5 only.
  - Microsoft Drive debug selftests reported `passed=28, failed=6`.
  - The failing cases covered descendant move rejection, empty-name child parsing, and a 70-level folder merge that should trip the provider recursion cap.

## GREEN

- Build: `.\build.ps1 -ProjectName FileSystemMicrosoftDrive -Configuration Debug`
  - Log: `GREEN-build-FileSystemMicrosoftDrive-msbuild-20260615_193548_402.log`
  - Result: 0 warnings, 0 errors.
- Build: `.\build.ps1 -ProjectName PluginContractTests -Configuration Debug`
  - Log: `GREEN-build-PluginContractTests-msbuild-20260615_193609_598.log`
  - Result: 0 warnings, 0 errors.
- Contract run: `.\.build\x64\Debug\PluginContractTests.exe`
  - Artifact: `GREEN-PluginContractTests.txt`
  - Result: passed.
  - Debug selftests: Microsoft Drive `passed=35, failed=0`; S3 `passed=102, failed=0`; curl `passed=68, failed=0`.

## Implementation Notes

- `MoveOrRenameItem(...)` now rejects same-path and descendant destinations with `ERROR_INVALID_PARAMETER` before Graph metadata/mutation calls.
- `ParseChildren(...)` skips empty-name entries so malformed Graph children cannot recurse on the same source/destination path.
- `MergeMoveFolderIntoExisting(...)` carries a recursion depth counter and returns `ERROR_STACK_OVERFLOW` at the provider cap (`64`) instead of walking unbounded.
