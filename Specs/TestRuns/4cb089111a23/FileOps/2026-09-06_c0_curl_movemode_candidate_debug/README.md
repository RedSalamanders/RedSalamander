# Curl Move-mode repair — Debug

Base `f0710f9ba47d3fc0a45413fffa3ffa84ac038074`; the exact eight-file candidate
patch and raw source hashes are retained here. Full test-enabled Debug receipt
`62993faa2bc62637fbbde4fdd3ef40a811f059a93e72b470d1a8db71e79eb720`, snapshot
`049d8724b6ba200a243b870ce24bd6edc7c1b1c32925b225e823518c415afa5f`.
Build: 184,614 ms, zero warnings/errors, including review of the raw build log.

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild -TestRoot C:\RedSalamander.Perf -CaseFilter C0_CurlMovePreflightTruth -RunId c0-curl-movemode-candidate-debug-20260906`.

Result: **247/247 native checks**, 3/0/0 suite, no skips. Native case: 20,234 ms.
All 48 mode rows and both mixed-bulk rows pass. Invalid options cause zero server
commands. Cross-endpoint native-only Move causes zero mutations/relay and reports
known no-commit. Same-endpoint file and nested/empty directory Moves still rename;
prefix siblings survive. Mixed batches at concurrency 1/4 produce partial aggregate
status, one native rename, and independently truthful committed/refused receipts.
They take 7,481 / 6,695 us, with two rename commands and 21 total commands each.

Same-machine Debug comparison to `../2026-09-06_c0_curl_movemode_baseline_debug/`:

| Equal file-scenario subset | Baseline / candidate duration sum (us) | Mutation commands | All commands |
|---|---:|---:|---:|
| Four valid native Moves | 17,242 / 15,366 | 8 / 8 | 60 / 60 |
| Four cross-endpoint native-only requests | 38,664 / 8,585 | 16 / 0 | 144 / 24 |
| Sixteen invalid/non-Move mode requests | 94,268 / 176 | 40 / 0 | 356 / 0 |

These are directional single-run timings. Refusal rows intentionally do less work;
they establish safe admission cost, not improved transfer throughput. The baseline
has 241 checks; this candidate adds six mixed-bulk assertions and strengthens the
directory corpus after repairing the fake server's file-only rename limitation.
Do not compare whole-suite duration as equal-work performance.

Validation reuses existing progress/admission and mutation-receipt owners, with no
new policy model. Native-only admission remains immutable across progress callbacks.
Default relay Move, whole-tree commitment memory, conditional source cleanup, and
live endpoint/session qualification remain separate contracts. This focused run
does not close C0 or replace broad Debug/Release or final Fresh Full.

Exact-copy and ordered-metric provenance accompany this archive. Build-log
normalization removes trailing horizontal whitespace only. Later plan/archive
documentation is not claimed as a new runtime result.
