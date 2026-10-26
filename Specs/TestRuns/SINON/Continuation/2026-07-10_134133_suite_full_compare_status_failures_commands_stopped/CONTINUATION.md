# Suite Full Continuation Snapshot

- Date: 2026-07-10 13:41 local
- Branch: codex/test-suite-stabilization-flake-convergence
- Suite command: `REDSALAMANDER_TEST_ROOT=C:\RSPerf .\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2`
- Run root: `C:\RSPerf\runs\20260710T112708Z-65924-0e5acf85d34443aabdefde3ac6f51018`
- Wrapper log base: `.build\codex-runs\suite-full-final-after-find-uia-pump-20260710_132708`

## Result

This is not accepted final evidence. The run built cleanly, then CompareDirectories failed two foreground search-service cases before the wrapper continued into Commands.

CompareDirectories:
- Passed: 217
- Failed: 2
- Skipped: 30
- Duration: 353974 ms

Failing cases:
- `search_service_compact_request_roundtrip`
  - Reason: timed out waiting for `SearchServiceBroker::GetStatus` for live compact request preflight service.
  - Expected pipe: `\\.\pipe\RedSalamander.SearchService.Test.A7B70266-002E-43A0-929A-CBDA93A6F2BF`
  - Last HRESULT: `0x80070002`
- `search_service_status_and_query_roundtrip`
  - Reason: timed out waiting for `SearchServiceBroker::GetStatus` for status roundtrip service.
  - Expected pipe: `\\.\pipe\RedSalamander.SearchService.Test.DFD5D827-943C-4B04-BF45-208B71613504`
  - Last HRESULT: `0x80070002`

Commands was still running when this snapshot was captured:
- Passed at capture: 340
- Failed at capture: 0
- Skipped at capture: 0
- Last observed cases were in the Preferences keyboard live interaction group.

## Captured Evidence

- `wrapper\suite-full-final-after-find-uia-pump-20260710_132708.out.log`
- `build\msbuild-20260710_132709_464.log`
- `compare\results.json`
- `compare\trace.txt`
- `compare\failures.json`
- `commands\results.json`
- `commands\trace.txt`
- `process\process-snapshot-before-stop.json`

## Next Step

Root-cause and fix the two foreground search-service status timeouts. Do not quarantine them and do not mark them flaky/non-blocking. After the fix, rerun focused cases, broad CompareDirectories, and then a fresh Suite Full from a clean `C:\RSPerf` root.
