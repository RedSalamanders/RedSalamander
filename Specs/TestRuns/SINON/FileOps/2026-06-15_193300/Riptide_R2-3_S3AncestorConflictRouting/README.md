# Riptide R2-3 S3 Ancestor Conflict Routing

Scope: prove S3 prefix copies over object-as-directory blockers use the per-child conflict pipeline, and that stacked blockers are re-probed after each removal.

## RED

- Build: `.\build.ps1 -ProjectName FileSystemS3 -Configuration Debug`
  - Log: `RED-build-FileSystemS3-msbuild-20260615_192812_359.log`
  - Result: 0 warnings, 0 errors. This build included the new in-memory S3 graph selftests but not the routing fix.
- Build: `.\build.ps1 -ProjectName PluginContractTests -Configuration Debug`
  - Log: `RED-build-PluginContractTests-msbuild-20260615_192912_615.log`
  - Result: 0 warnings, 0 errors.
- Contract run: `.\.build\x64\Debug\PluginContractTests.exe`
  - Result: failed for the intended reason in Step5 only.
  - S3 debug selftests reported `passed=91, failed=8`.
  - The failing cases covered prefix-over-root-object Skip/Overwrite and a stacked `dest/a` plus `dest/a/b` object-as-directory blocker.

## GREEN

- Build: `.\build.ps1 -ProjectName FileSystemS3 -Configuration Debug`
  - Log: `GREEN-build-FileSystemS3-msbuild-20260615_193114_397.log`
  - Result: 0 warnings, 0 errors.
- Build: `.\build.ps1 -ProjectName PluginContractTests -Configuration Debug`
  - Log: `GREEN-build-PluginContractTests-msbuild-20260615_193227_572.log`
  - Result: 0 warnings, 0 errors.
- Contract run: `.\.build\x64\Debug\PluginContractTests.exe`
  - Artifact: `GREEN-PluginContractTests.txt`
  - Result: passed.
  - Debug selftests: Microsoft Drive `passed=25, failed=0`; S3 `passed=102, failed=0`; curl `passed=68, failed=0`.

## Implementation Notes

- `ExecuteCopyOrMove(...)` no longer rejects `source prefix -> destination object` before the per-child conflict pipeline. The destination object is now discovered as an ancestor blocker for the generated child key.
- `RefreshDestinationState(...)` centralizes exact-object and ancestor-object probes so conflict state is refreshed after mutations instead of trusting stale precomputed blockers.
- The per-object loop now backs up and removes one ancestor object at a time, then re-probes. Callback-driven overwrite is reset to the global overwrite grant after each ancestor removal so newly discovered blockers are prompted independently.
- If a later newly discovered blocker is skipped after an earlier ancestor was already backed up/deleted for the same child, `RestoreBackupsFrom(...)` restores that per-object work before continuing with a partial result.
