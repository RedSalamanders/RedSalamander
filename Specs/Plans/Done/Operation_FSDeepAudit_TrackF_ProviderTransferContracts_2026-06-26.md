# Operation FS Deep Audit Track F - Provider Transfer, Identity, and Delete-Proof Contracts - 2026-06-26

## Status

- State: Done 2026-06-27. Ready to move to `Specs/Plans/Done/`.
- Priority: P1 provider data safety.
- Scope: 7z, Curl/IMAP, S3, Microsoft Drive, Google Drive, and Dummy FS contract gaps.

## Problem

The generic cross-FS bridge now rejects unverified unknown-size transfers, but provider bodies can still produce short reads, weak delete proof, inconsistent overwrite behavior, or identity ambiguity. Coverage must use fake transports/providers before relying on live services.

## Targets

- `Plugins/FileSystem7z/*`
- `Plugins/FileSystemCurl/*`
- `Plugins/FileSystemS3/*`
- `Plugins/FileSystemMicrosoftDrive/*`
- `Plugins/FileSystemGoogleDrive/*`
- `Plugins/FileSystemDummy/*`
- `Specs/Plugins/Plugins_VirtualFileSystem.md`

## Tasks

1. [x] 7z: unsafe archive keys are rejected during normalization (`//`, drive-qualified, empty component, `.`, `..`, colon/ADS-like component, embedded NUL) and file/directory or duplicate file key collisions now fail indexing with `ERROR_INVALID_DATA`.
2. [x] 7z: short extraction/spool reads were already fail-closed in `SevenZipItemFileReader` (`ERROR_HANDLE_EOF`/terminal extract status). No additional code change was needed after review.
3. [x] Curl/HTTP/FTP staged upload: existing implementation probes the staged remote size, falls back to entry metadata when needed, deletes the staged object on mismatch, and returns `ERROR_PARTIAL_COPY`. Source-contract coverage was updated to pin that behavior.
4. [x] Curl/IMAP: IMAP same-provider move/rename are not enabled in the current capability surface; the remaining UIDVALIDITY/EXPUNGE concerns are not active mutation paths for this shipped contract.
5. [x] S3: ranged reads now prove known size first and validate expected range length against response `Content-Length` and actual body bytes. Short, overlong, and negative-length responses fail.
6. [x] S3: delete-after-copy, pagination, ancestor collision, and backup/orphan cleanup contracts were already covered by fake S3 graph debug self-tests and Floodgate source-contract guards; new ranged-read debug self-tests were added.
7. [x] Microsoft Drive: upload-session/content-range, provider-id, invalid-options, empty-child, merge-move re-list, and backup cleanup behavior already had fake Graph debug self-tests and capability/source-contract guards.
8. [x] Google Drive: same-provider mutation APIs remain fail-closed/unsupported and capabilities advertise `pathTextStableIdentity = false`; duplicate/unnamed display names are therefore not active destructive or overwrite paths.
9. [x] Dummy FS and host guards remain covered by plugin capability parsing, debug self-tests, and source-contract checks.
10. [x] `Specs/Plugins/Plugins_VirtualFileSystem.md` now carries durable stream, archive-key, staged upload, provider proof, and unstable-identity mutation contracts.

## Implementation Notes

- `Plugins/FileSystem7z/FileSystem7z.cpp`
  - `NormalizeArchiveEntryKey` now rejects unsafe absolute, drive-qualified, traversal, stream/ADS-like, and malformed component keys.
  - `BuildIndex` fails duplicate file keys and file/directory key collisions instead of silently overwriting `outEntries[raw.key]`.
- `Plugins/FileSystemS3/FileSystemS3.IO.cpp`
  - Added `FsS3::ValidateS3RangeResponseLength(...)` and wired `S3RangedFileReader` through it.
  - Added `FsS3::RunDebugRangeReadContractSelfTest(...)` and exported it through `RedSalamanderS3DebugSelfTests`.
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`
  - Updated Floodgate provider-transfer source guards for the current S3/Curl/7z/MSDrive implementations.

## Validation

```powershell
.\build.ps1 -ProjectName FileSystemS3
.\build.ps1 -ProjectName FileSystem7z
Invoke-Pester -Path Tools\Tests\TestHarnessSourceContracts.Tests.ps1
Push-Location .build\x64\Debug; .\PluginContractTests.exe; Pop-Location
```

2026-06-27 results:
- `FileSystemS3` Debug build passed (`.build\logs\msbuild-20260627_112841_190.log`).
- `FileSystem7z` Debug build passed (`.build\logs\msbuild-20260627_113142_832.log`).
- Source-contract Pester suite passed: 30 passed / 0 failed.
- `PluginContractTests.exe` passed; `FileSystemS3.dll` debug self-tests passed 115 / 0 failed.
