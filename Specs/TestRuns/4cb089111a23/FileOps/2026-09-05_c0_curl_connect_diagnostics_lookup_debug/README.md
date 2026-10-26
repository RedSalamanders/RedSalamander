# C0 Curl connection-failure diagnostics and continuation

Based on `4a42e3cbba30bedbf0426296f0a0b20a97b19b11`.
Final Debug/Release builds bind source snapshot
`2afb2008ab09ed5d40e950d4c6020f36f04e328c52663586e86e56c3ce70e428`.

The diagnostic step is validated; **C0 is still open**. The broad failure now has
an underlying Windows code: **10048 (local address/port already in use)**, not
10061 (refused connection). The old HRESULT mapping conflates those causes and
is unchanged here. No connection-lifecycle repair is claimed.

## Implementation and evidence

Cursor completion, each size-probe attempt and single namespace quote dispatch
capture Curl/OS codes before handle reset. The shared transport-local helper
retains numeric phase/availability/ports/reply-code data only, never URL/profile,
credentials, control reply text or the Curl error buffer. Positive ports can
describe the control connection, not the failed data socket; unavailable values
stay explicit. Success/cancel paths do not perform the diagnostic queries.

The refusal control owns an exclusively bound, non-listening loopback socket for
its entire lifetime. Another server cannot acquire the test port. The three calls
must fail and publish no entry/known size; the namespace call must not succeed.
Actual HRESULTs and call counts remain visible. Native exact-count guards expect
237 lookup / 142 Copy checks.

| Run | Native pass / fail / skip | Evidence |
|---|---:|---|
| [Initial Release control](../2026-09-05_c0_curl_connect_diagnostics_initial_release/README.md) | 2 / 1 / 0 | Lookup 234/237; three new operation checks fail with the too-short initial control deadline |
| [Broad Release C0](../2026-09-05_c0_curl_connect_diagnostics_broad_release/README.md) | 15 / 1 / 1 | Copy 130/142; eight cursor failures carry Windows 10048 |
| [Final Release lookup](../2026-09-05_c0_curl_connect_diagnostics_lookup_release/README.md) | 3 / 0 / 0 | Lookup 237/237; all five diagnostic rows carry 10061 |
| This Debug lookup | 3 / 0 / 0 | Lookup 237/237; all five diagnostic rows carry 10061 |
| [Release CRUD](../2026-09-05_c0_curl_connect_diagnostics_crud_release/README.md) | 3 / 0 / 0 | Successful paths remain green |
| [Debug CRUD](../2026-09-05_c0_curl_connect_diagnostics_crud_debug/README.md) | 3 / 0 / 0 | Successful paths remain green |

Native exact-case counts include Setup/Cleanup. Debug suite time is 20,281 ms.
The refusal control takes 11.204864 s and reports exactly one cursor attempt,
three size-probe attempts and one namespace attempt; every actual HRESULT is
`0x800704C9`, and the control verdict is S_OK. Release shows the same counts.
The host trace confirms all 237 checks; expected negative metric HRESULTs are
deliberate inputs, not failed native verdicts.

The initial 500 ms connect limit was shorter than an independent owned-socket
probe's observed refusal latency (2,064 ms). The corrected control uses 5 s connect /
7 s operation limits. Production/traversal limits, 4,097-file corpus, 120-second
deadline, 64-creation cap, full-list validation, mutation/retry rules and fake
listing generation are unchanged. This is not a throughput optimization.

## Build and policy provenance

Final test-enabled full-solution builds have zero warnings/errors:

- Debug: 03:11, receipt `2b2adacb8a58d3313c4722645f5c09a568baac2b7f0e3ef05d4c0fb106b5a9ad`.
- Release: 04:30, receipt `10e81ea099b865401a5a72d83377a42534d8087a3dc6011a48c9cbabe17eaf52`.

An attempted overlapping Debug build was rejected before MSBuild by the shared
dependency lock owned by the Release build (owner PID 121092, platform X64).
It was retried only after Release completed; no process or lock was removed.
The initial Release build's MSB3061 shared-libcrypto Clean warning remains in the
initial archive. A later warning-free rebuild does not fix/prove C6's Clean ownership.

Final source/resource policy: **197 passed / 0 failed / 0 skipped**, 17.6866828 s;
the original XML from `c0-curl-connect-diagnostics-final-policy-20260905` is retained.
Initial policy XML is retained with the initial Release control. Scoped archive
attributes preserve exact raw bytes under autocrlf. Final archive/status/attribute
edits occur after runtime validation; tested C++/resource bytes do not change.
This is focused evidence, not a final-HEAD Fresh Full claim.

## Next bounded step

1. Use the captured 10048 to trace owned loopback control/data socket lifetimes and
   port allocation pressure in the broad process. Do not infer exhaustion or the
   owner from machine-wide TIME_WAIT counts alone. Do not widen the OS port range,
   reduce TIME_WAIT, add resets/reuse tuning, raise traversal timeouts, split/reorder
   the suite or suppress checks just to turn it green.
2. After a causal repair, repeat the predecessor cluster in one process:
   `C0_CurlDeleteTraversalBounds,C0_CurlDirectorySizeTruth,C0_CurlMovePreflightTruth,C0_CurlCopyTraversalTruth`,
   then `C0_`. Use the governed runner and fresh unique IDs. An isolated Copy pass
   is already known and insufficient.
3. Separately optimize the measured fake `BuildDirectoryPayload` allocation/
   serialization cost, with literal-byte/fault oracles before changing generation.
   Previous Debug serial wide Copy spends 53.06 s in nested listing generation/
   mutex wait and times out at 3,675/4,097 files. Preserve full-list validation and
   source-size proof; distinguish fixture speed from real-provider performance.
4. Re-run lookup, broad C0 and CRUD in Debug/Release, archive failures and exact
   receipts, validate policy/spec/archive hygiene and commit. Then continue C0's
   Move commitment memory, IMAP, real FTP/SSH quotas, route/walker and host Rename
   work, followed by C1–C6. Final Fresh Full and inherited I17 remain required.

The Hello exception stays confined to the five previously approved isolated
profiles. No normal settings, TLS/SSH policy or configured-server data changed.
Optional questions about verified-TLS IMAP and unique-child live mutations remain
unanswered; do not infer either authorization. The user-owned repository-root
`last_run/` is untouched. Retained sandbox runs/disk-audit findings are not deleted.

Microsoft distinguishes [Winsock 10048 from refusal](https://learn.microsoft.com/en-us/windows/win32/winsock/windows-sockets-error-codes-2).
Curl documents the [underlying OS error](https://curl.se/libcurl/c/CURLINFO_OS_ERRNO.html)
and [last connection's port metadata](https://curl.se/libcurl/c/CURLINFO_PRIMARY_PORT.html).
These references explain interpretation; archived native measurements establish
what this run actually observed.

## Reproduction and integrity

Run `c0-curl-connect-diagnostics-final-lookup-debug-20260905`:

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter C0_CurlEntryLookupTruth -RunId c0-curl-connect-diagnostics-final-lookup-debug-20260905
```

Raw evidence is retained in the named governed run under `C:/RedSalamander.Perf/runs/`. Native/root/runner results, traces and entry fingerprints are copied byte-for-byte, with hashes in `file-provenance.json`. `build-provenance.json` binds the test-enabled full-solution receipt. No settings or credentials are archived. Use a new run ID for a repeat.
