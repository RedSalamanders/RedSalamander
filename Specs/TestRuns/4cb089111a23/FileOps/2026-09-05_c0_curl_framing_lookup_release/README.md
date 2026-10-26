# Curl framing: Release lookup correctness

Plugin lookup passes 233/233; native/runner 3 passed, 0 failed, 0 skipped.
Suite 7,906 ms, named case 6,485 ms. All 64 route rows, 32 framing rows, four
wide repeated-list probes and eight cancellation/reborrow rows pass. All 443 raw
metric rows remain unchanged. The test-enabled full-solution Release build took
04:29 with zero warnings/errors and binds the same source snapshot as final Debug.

The initial Release build attempt stopped before compilation because its process
ownership check could not inspect residual tools (Access denied). The approved
retry used the existing guard unchanged. No process was killed, no lock bypassed.

## Reproduction and provenance

Run `c0-curl-framing-final-lookup-release-20260905` on machine `4cb089111a23`.

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter C0_CurlEntryLookupTruth -RunId c0-curl-framing-final-lookup-release-20260905
```

Use a new run ID when repeating. `build-provenance.json` and the unchanged entry
fingerprints bind source/build identities; `file-provenance.json` records raw hashes.
Results and traces are unchanged. If present, `build.txt` only normalizes trailing
horizontal whitespace and UTF-8 encoding; both log hashes are recorded and the raw
log remains under `.build/logs`. Raw runtime data remain in this exact run beneath
`C:\RedSalamander.Perf`. No retained data were deleted to clear disk-audit findings.

See the [step overview](../2026-09-05_c0_curl_framing_lookup_debug/README.md).
This is focused evidence, not final Fresh Full or plan retirement.
