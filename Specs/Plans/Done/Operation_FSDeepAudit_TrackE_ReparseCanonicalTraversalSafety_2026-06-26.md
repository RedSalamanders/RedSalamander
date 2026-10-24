# Operation FS Deep Audit Track E - Reparse, Canonicalization, and Traversal Safety - 2026-06-26

## Status

- State: Complete; ready to move to `Specs/Plans/Done/`.
- Priority: P1 data safety and traversal containment.
- Scope: Local reparse behavior, path canonicalization, watch/search traversal, and cycle guards.

## Problem

Metadata, delete, watch, search, and traversal paths must consistently distinguish link-object operations from link-target operations. Extended path forms and recursive traversal also need explicit containment and cycle guards.

## Targets

- `Plugins/FileSystem/FileSystem.Path.cpp`
- `Plugins/FileSystem/FileSystem.FileOps.cpp`
- `Plugins/FileSystem/FileSystem.Watch.cpp`
- `Plugins/FileSystem/FileSystem.Search.cpp`
- `Common/LocalSearchIndexCore.cpp`
- `Specs/FileSystem/FileSystem_FileOperations.md`
- `Specs/Core/Core_Search.md`

## Tasks

1. [x] Classify every local open by intent: target contents, object metadata, reparse object, or recursive traversal.
2. [x] Add `FILE_FLAG_OPEN_REPARSE_POINT` where metadata/watch/delete must apply to the link object.
3. [x] Normalize `\\?\` and normal Win32 paths consistently, including `.`/`..` collapse for supported extended drive and UNC forms.
4. [x] Preserve display casing separately from comparison keys.
5. [x] Add visited-set/depth/queue guards for recursive traversal, with visible incomplete-result warnings.
6. [x] Add tests for metadata-on-link, `\\?\` canonicalization/root-escape behavior, junction cycle traversal, and watch setup behavior on reparse roots.
7. [x] Update File Operations and Core Search specs.

## Closeout Notes

- Existing implementation already covered the reparse-object destructive contract: local delete and overwrite cleanup open no-follow with `FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS`; recursive delete re-opens queued children before destructive recursion decisions; link/reparse properties expose link-object target metadata without placeholder rows.
- Existing search traversal already kept native follow-symlink traversal bounded by folded logical path plus physical identity, queued identity caps, access-denied skips, and overflow warnings.
- The meaningful remaining gap was path canonicalization: `MakeAbsolutePath(...)` returned supported `\\?\` paths unchanged, and `ToExtendedPath(...)` could bypass canonicalization for already-extended paths. Fixed in `Plugins/FileSystem/FileSystem.Path.cpp` by stripping/restoring only supported extended drive and UNC prefixes around the Win32 full-path resolver. Unsupported device namespaces remain unchanged and are not silently converted into ordinary local filesystem paths.
- Added `RunDebugPathNormalizationSelfTest(...)` and wired it into `RedSalamanderFileSystemDebugSelfTests`, covering ordinary drive-rooted paths, drive-rooted extended paths, extended UNC paths, and unsupported device namespace preservation.
- Updated `Specs/FileSystem/FileSystem_FileOperations.md` and `Specs/Core/Core_Search.md` with the durable canonical local path contract and reparse-open intent rules.

## Validation

```powershell
.\build.ps1 -ProjectName FileSystem
# PASS, 0 warnings, 0 errors
# Log: .build\logs\msbuild-20260627_112445_679.log

Push-Location .build\x64\Debug
.\PluginContractTests.exe
Pop-Location
# PASS; FileSystem.dll debug selftests passed=35 failed=0, including path normalization selftest.
```
