# Operation FS Deep Audit Track D - Search Service Trust Boundary and Pipe Protocol - 2026-06-26

## Status

- State: Complete; moved to Done on 2026-06-27.
- Priority: P1 security and service reliability.
- Scope: Search service authorization, named-pipe framing, cancellation, and shutdown semantics.

## Problem

The search service can run with service authority while accepting pipe requests that carry client paths and search criteria. The service must prove the client is allowed to access requested roots, and the pipe protocol must fail short/slow/malformed frames without hanging the service.

## Targets

- `Common/SearchServiceBroker.cpp`
- `RedSalamanderSearchService/Main.cpp`
- `Specs/Core/Core_Search.md`
- Search/Compare service selftests or a dedicated fake pipe harness

## Tasks

1. Done: documented service identity, local pipe trust boundary, request-root authorization, and disconnect-and-continue protocol rules in `Specs/Core/Core_Search.md`.
2. Done: service query/rebuild roots are normalized and authorized under `ImpersonateNamedPipeClient(...)` before service authority opens the requested root.
3. Done: pipe remains local-only with `PIPE_REJECT_REMOTE_CLIENTS` and the documented SDDL contract (`SY`/`BA` full access, interactive clients read/write).
4. Done: service rejects empty, malformed, relative, UNC, `\\?\UNC`, `\\?\GLOBALROOT`, `\\.\`, and `\??\` roots before query/rebuild dispatch.
5. Done: server-side pipe reads/writes now use overlapped I/O with per-frame deadlines and the service stop event.
6. Done: server frame receive/send paths bound frame size, detect zero-byte/short transfer failure, and return protocol failures instead of blocking indefinitely.
7. Done: recoverable client failures are disconnected while the server continues accepting later clients; outstanding server I/O is cancelled during timeout/shutdown.
8. Done: added `search_service_rejects_device_root_and_continues` and `search_service_slow_partial_client_does_not_block_next_client`, plus static case inventory entries.
9. Done: updated `Specs/Core/Core_Search.md` diagnostics and normative summary.

## Implementation Notes

- `Common/SearchServiceBroker.cpp` now carries a `ServerIoContext` through service frame send/receive paths.
- `RunServer(...)` applies an effective default storage root before bootstrap/status handling and classifies protocol, access, malformed-root, timeout, and disconnect HRESULTs as recoverable client-scoped failures.
- Query and rebuild handling canonicalizes the client root and impersonates the pipe client before the repository sees the root.
- The selftest runner requires static compare-case inventory updates in `CompareDirectoriesEngine.SelfTest.cpp`; the two new cases are registered there as well as in the `RunCase(...)` source.

## Validation

```powershell
.\build.ps1 -ProjectName RedSalamander -Rebuild
# PASS, 0 warnings, 0 errors
# Log: .build\logs\msbuild-20260627_105735_015.log

.\build.ps1 -ProjectName RedSalamander
# PASS, 0 warnings, 0 errors
# Log: .build\logs\msbuild-20260627_110630_620.log

$env:REDSALAMANDER_SELFTEST_ROOT = 'D:\RedSalamander\Specs\TestRuns\Search\2026-06-27_TrackD_SearchServiceTrustBoundaryPipeProtocol'
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter 'search_service_rejects_device_root_and_continues,search_service_slow_partial_client_does_not_block_next_client'
# PASS: 2 passed, 0 failed

$env:REDSALAMANDER_SELFTEST_ROOT = 'D:\RedSalamander\Specs\TestRuns\Search\TDSvc'
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter 'search_service_'
# PASS: 30 passed, 0 failed, 4 skipped for machine-dependent preconditions
```

Note: longer selftest artifact roots can push default ProgramData SQLite test paths over legacy Win32 path limits. The broad service sweep above uses a short durable artifact root for meaningful default-store coverage.
