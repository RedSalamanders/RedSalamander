# Riptide R2-5 Storage Probe Volume Resolution

Scope: prove local storage probing resolves mounted-volume and SUBST-style paths to the backing volume device before `IOCTL_STORAGE_QUERY_PROPERTY`, instead of opening the lexical drive-letter device.

## RED

- Build: `.\build.ps1 -ProjectName FileSystem -Configuration Debug`
  - Log: `RED-build-FileSystem-msbuild-20260615_194205_364.log`
  - Result: 0 warnings, 0 errors. This build included the new debug selftest/export and the probe-open seam, but not the volume-root/GUID/SUBST resolver.
- Build: `.\build.ps1 -ProjectName PluginContractTests -Configuration Debug`
  - Log: `RED-build-PluginContractTests-msbuild-20260615_194234_223.log`
  - Result: 0 warnings, 0 errors.
- Contract run: `.\.build\x64\Debug\PluginContractTests.exe`
  - Result: failed for the intended reason in Step5 only.
  - `FileSystem.dll` debug selftests reported `passed=2, failed=5`, proving the old probe still opened `\\.\R:` and did not resolve the mounted volume root or volume GUID.

## GREEN

- Build: `.\build.ps1 -ProjectName FileSystem -Configuration Debug`
  - Log: `GREEN-build-FileSystem-msbuild-20260615_194743_690.log`
  - Result: 0 warnings, 0 errors.
- Build: `.\build.ps1 -ProjectName PluginContractTests -Configuration Debug`
  - Log: `GREEN-build-PluginContractTests-msbuild-20260615_194519_565.log`
  - Result: 0 warnings, 0 errors.
- Contract run: `.\.build\x64\Debug\PluginContractTests.exe`
  - Artifact: `GREEN-PluginContractTests.txt`
  - Result: passed.
  - Debug selftests: FileSystem `passed=14, failed=0`; Microsoft Drive `passed=35, failed=0`; S3 `passed=102, failed=0`; curl `passed=68, failed=0`.

## Implementation Notes

- `GetStorageCharacteristics(...)` now resolves the real local volume root with `GetVolumePathNameW(...)` before probing storage media.
- `ProbeLocalStorageMedium(...)` opens the resolved `\\?\Volume{GUID}` device from `GetVolumeNameForVolumeMountPointW(...)`; the old `\\.\<letter>:` path remains only as a fallback.
- SUBST-style drive aliases fall back through `QueryDosDeviceW("<letter>:")` and then re-resolve the backing target's volume root/GUID before the IOCTL path is opened.
- The same resolved volume root is used for `GetDriveTypeW(...)` in transfer/storage hint entrypoints when available, so mounted volumes do not inherit the host letter's drive type.
