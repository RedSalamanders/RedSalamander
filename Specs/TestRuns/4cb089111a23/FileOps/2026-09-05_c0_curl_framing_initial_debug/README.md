# Curl framing: initial candidate failure retained

The new production framing and all 32 added checks passed: plugin 233 passed,
0 failed, S_OK. The host still expected 201, so native/runner correctly failed:
2 passed, 1 failed, 0 skipped, suite 9,360 ms. The build also emitted one C5246
missing-braces warning (03:00, zero errors). Neither failure is reclassified here.

The corrected candidate adds the required `std::array` aggregate braces and
synchronizes exact host expectations to lookup 233 and Copy 142. Copy's stale
141 expectation was inherited from the preceding instrumentation step. The guard
still requires equality; its diagnostic now reports observed and expected counts.
All 440 raw metric rows and the initial 197/197 policy result remain intact.

## Reproduction and provenance

Run `c0-curl-framing-candidate-debug-20260905` on machine `4cb089111a23`.

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -TestRoot 'C:\RedSalamander.Perf' -CaseFilter C0_CurlEntryLookupTruth -RunId c0-curl-framing-candidate-debug-20260905
```

Use a new run ID when repeating. `build-provenance.json` and the unchanged entry
fingerprints bind source/build identities; `file-provenance.json` records raw hashes.
Results and traces are unchanged. If present, `build.txt` only normalizes trailing
horizontal whitespace and UTF-8 encoding; both log hashes are recorded and the raw
log remains under `.build/logs`. Raw runtime data remain in this exact run beneath
`C:\RedSalamander.Perf`. No retained data were deleted to clear disk-audit findings.

See the [step overview](../2026-09-05_c0_curl_framing_lookup_debug/README.md).
This is focused evidence, not final Fresh Full or plan retirement.
