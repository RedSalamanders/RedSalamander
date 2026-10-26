# Curl framing: isolated Release Copy diagnostic

Same source and Release receipt as the failed broad run, with no intervening code
change: plugin 142/142, all 46 scenarios, native/runner 3 passed, 0 failed,
0 skipped. Suite 55,484 ms; named case 54,094 ms. Both wide copies publish all
4,097 exact files: parallel 7,158,096 us and serial 24,593,820 us. All depth-128
Move and depth-129 refusal checks now pass. All 970 raw metric rows remain intact.

This proves the exact case can pass, not that broad C0 is fixed. Keep the failed
[broad result](../2026-09-05_c0_curl_framing_wide_release/README.md) open.
Read-only system loopback connection queries were made during/around this diagnostic
run, so do not use its timings as a clean performance baseline.

A query returned 35,179 loopback TIME_WAIT entries at observation timestamp
2026-09-05T19:57:58.5421846Z. An earlier query returned 690; the configured IPv4 TCP
dynamic range was 49,152 + 16,384 ports. These machine-wide, non-atomic observations
are not owned-socket counts and do not prove exhaustion or the cause of a refusal.
No network setting, timeout, reuse option, connection, or external process was changed.
Next instrumentation must capture the failing Curl result, OS error and owning
route/session lifecycle before attributing the broad failure to resource pressure.

## Reproduction and provenance

Run `c0-curl-framing-copy-isolated-release-20260905` on machine `4cb089111a23`.

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter C0_CurlCopyTraversalTruth -RunId c0-curl-framing-copy-isolated-release-20260905
```

Use a new run ID when repeating. `build-provenance.json` and the unchanged entry
fingerprints bind source/build identities; `file-provenance.json` records raw hashes.
Results and traces are unchanged. If present, `build.txt` only normalizes trailing
horizontal whitespace and UTF-8 encoding; both log hashes are recorded and the raw
log remains under `.build/logs`. Raw runtime data remain in this exact run beneath
`C:\RedSalamander.Perf`. No retained data were deleted to clear disk-audit findings.

See the [step overview](../2026-09-05_c0_curl_framing_lookup_debug/README.md).
This is focused evidence, not final Fresh Full or plan retirement.
