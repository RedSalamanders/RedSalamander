# C0 Curl bounded framing and exact test-count guards

Based on `809daa66dfb6179b7ad83fb7ec38412f1689361f`. Final Debug and test-enabled
Release bind source snapshot
`673426f0da99bcb2f72a90557ff98fa35f8d7574212dc5fbf31bbae243ac4530`.

The cursor now scans/copies bounded contiguous fragments, preserving the strict
parser, two buffers, pool, cancellation checkpoints, successful EOF requirement
and memory reservation. The 32 new deterministic callback/consumer checks cover
all chunk splits, UTF-8/CRLF, exact record limits, overflow, NUL and late malformed
rows. This seam does not replace transport/cancellation/mutation route tests.

The host guard now expects 233 lookup and 142 Copy checks and reports both
expected and observed counts. It still fails on a mismatch; no check is dropped.
The initial warning and stale-count failure are preserved below.

## Evidence, including failures

| Run | Native pass / fail / skip | Plugin / result |
|---|---:|---|
| [Unchanged Debug baseline](../2026-09-05_c0_curl_framing_baseline_debug/README.md) | 3 / 0 / 0 | Lookup 201/201 |
| [Initial Debug candidate](../2026-09-05_c0_curl_framing_initial_debug/README.md) | 2 / 1 / 0 | All 233 plugin checks pass, host expected 201; one C5246 warning |
| This final Debug lookup | 3 / 0 / 0 | Lookup 233/233 |
| [Broad Debug C0](../2026-09-05_c0_curl_framing_wide_debug/README.md) | 15 / 1 / 1 | Copy 141/142, serial wide timeout |
| [Debug basic CRUD](../2026-09-05_c0_curl_framing_crud_debug/README.md) | 3 / 0 / 0 | Passed |
| [Release lookup](../2026-09-05_c0_curl_framing_lookup_release/README.md) | 3 / 0 / 0 | Lookup 233/233 |
| [Broad Release C0](../2026-09-05_c0_curl_framing_wide_release/README.md) | 15 / 1 / 1 | Copy 131/142; wide complete, later deep-tree failures |
| [Release basic CRUD](../2026-09-05_c0_curl_framing_crud_release/README.md) | 3 / 0 / 0 | Passed |
| [Isolated Release Copy diagnostic](../2026-09-05_c0_curl_framing_copy_isolated_release/README.md) | 3 / 0 / 0 | Copy 142/142; does not replace broad failure |

Every exact-case native count includes Setup and Cleanup. The broad skipped case
is lookup, not reached after Copy failure despite `fail_fast=false`. Both final
lookup configurations retain 64 route, 32 framing, four repeated-list, and eight
cancel/reborrow observations, all passing. Final source/resource policy is 197/197,
zero skipped; its original NUnit XML is retained here.

Debug full-solution build: 03:13, zero warnings/errors, receipt
`b904205f914805316ef312e215e07e3a311516a952e77db95c6d1cf3f259ac05`.
Release full-solution build: 04:29, zero warnings/errors, receipt
`06203c1748576f7b6f57bc3bd52d50dec3e78b1131180b5eed9cddd187ecff1d`.
Both have `tests_enabled=true`. The focused runs reused verified receipts, not
unattested executables. The final lookup's suite time is 9,312 ms (case 7,844 ms).

## Performance interpretation

Same-machine 4,097-row, 16-query fixed-payload probes, microseconds:

| Probe | Unchanged Debug | Initial candidate | Final candidate |
|---|---:|---:|---:|
| Unix pure | 321,761 | 296,535 | 326,506 |
| DOS pure | 298,741 | 285,804 | 301,013 |
| Unix mixed | 340,228 | 289,071 | 307,730 |
| DOS mixed | 303,695 | 264,408 | 307,775 |

These are single-run samples: variation is material, so no consistent throughput
improvement is claimed. The final mixed-Unix probe is faster, the other three
slightly slower. New `.Framing` rows time correctness checks including the adapter;
they are not network performance measurements.

The unchanged wide-copy corpus, 120-second deadline, 64-listener cap, complete
listing/collision validation and exact contents remain enforced. Debug parallel
passes at 107.38 s; serial times out with 3,675/4,097 publications. Release wide
Copy completes in both modes, but the subsequent depth-128/129 integration failures
keep that broad run red. No fake-server generation was changed in this step, and
no real-server throughput or whole-operation memory bound is proved.

## Continue from here

1. Diagnose the broad Release deep-tree connection refusals with bounded,
   secret-free transport/OS-error and session-lifecycle evidence. Replay the
   predecessor cluster and broad C0 after a causal repair; an isolated pass is
   insufficient. No socket tuning, timeout increases or test suppression.
2. Separately reduce measured fake-listing serialization cost without changing
   emitted bytes, late faults, corpus, source-size proof or production validation.
   Add literal-payload oracles; distinguish fixture speed from product throughput.
   Current Debug serial generation/mutex wait is 53.06 s inside staging/promotion.
3. Repeat affected lookup/C0/CRUD in Debug and test-enabled Release, preserve every
   failure, and commit the next bounded correction. C0 still owns Move commitment
   memory, IMAP, real FTP/SSH limits, route/walker matrix and host Rename.
4. Continue C1–C6; no final Fresh Full, inherited I17 or plan retirement is claimed.

The Hello exception remains limited to five previously approved isolated profiles.
These runs use loopback fixtures, not new configured-server mutations. Pending
questions about isolated IMAP certificate verification and fresh-child-only live
mutations remain unanswered; normal settings and TLS/SSH policies were not changed.

## Reproduction and archive integrity

Run `c0-curl-framing-final-lookup-debug-20260905`:

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter C0_CurlEntryLookupTruth -RunId c0-curl-framing-final-lookup-debug-20260905
```

Use a new run ID when repeating. All 440 raw metric rows, native/root/runner
results, traces and entry fingerprints are unchanged. The final NUnit result is
from `c0-curl-framing-final-source-policy-20260905`. Each sibling archive records
its command, counts and limits. Large broad-run metrics use verified digest-bound
row projections; raw streams remain in their exact governed runs. Build logs only
normalize trailing horizontal whitespace and UTF-8 encoding, with both hashes in
`build-provenance.json`; the raw logs remain under `.build/logs`.
`file-provenance.json` binds all other raw copies. No secrets/settings are archived,
and no retained runs or the user-owned repository-root `last_run/` were removed.
The scoped `.gitattributes` rule preserves these nine archives' exact bytes under
`core.autocrlf`; it does not renormalize older evidence or other source files.
It and the final evidence/status prose were added after runtime validation; the
tested C++/resource bytes did not change. This is not a final-HEAD Fresh Full claim.
