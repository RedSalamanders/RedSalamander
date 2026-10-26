# Curl framing: unchanged Debug baseline

Unchanged `809daa66`: `C0_CurlEntryLookupTruth` passed 201/201 plugin checks;
native/runner 3 passed, 0 failed, 0 skipped (Setup and Cleanup included).
Suite 9,500 ms; build 03:05, zero warnings/errors. This application-target receipt
cannot authorize `-SkipBuild`. All 408 raw metric rows are retained unchanged.

## Reproduction and provenance

Run `c0-curl-framing-baseline-debug-20260905` on machine `4cb089111a23`.

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -TestRoot 'C:\RedSalamander.Perf' -CaseFilter C0_CurlEntryLookupTruth -RunId c0-curl-framing-baseline-debug-20260905
```

Use a new run ID when repeating. `build-provenance.json` and the unchanged entry
fingerprints bind source/build identities; `file-provenance.json` records raw hashes.
Results and traces are unchanged. If present, `build.txt` only normalizes trailing
horizontal whitespace and UTF-8 encoding; both log hashes are recorded and the raw
log remains under `.build/logs`. Raw runtime data remain in this exact run beneath
`C:\RedSalamander.Perf`. No retained data were deleted to clear disk-audit findings.

See the [step overview](../2026-09-05_c0_curl_framing_lookup_debug/README.md).
This is focused evidence, not final Fresh Full or plan retirement.
