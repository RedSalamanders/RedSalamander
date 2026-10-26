# R4-A19 unchanged-scheduler provider and independent-volume qualification

- Machine profile: `4cb089111a23`
- Observation implementation: `0c7ac7a4`
- Parent matrix implementation: `dae2f674`
- Final cancellation-result test correction: `22892314`
- Test-enabled x64 Release full-solution receipt: `83b15f05880242a2dda22bdbfd10f788c1de15f4181f33559ac0bcda4fa1e99c`; diagnostics `0 warning(s), 0 error(s)`
- Command: `.build\x64\Release\RedSalamander.exe --fileops-selftest --selftest-no-activate --selftest-timeout-multiplier=4 --selftest-case=FileOpsFamily_R4A19DiscoveryBaseline`
- Environment: `REDSALAMANDER_TEST_ROOT=C:\RedSalamander.Perf`, `REDSALAMANDER_TEST_ALLOW_ALTERNATE_VOLUME=1`, and one unique `REDSALAMANDER_TEST_RUN_ID`/process-local TEMP root per process
- Raw archives: adjacent `2026-09-02_r4_a19_parent_release_01` through `_05`

Each independent process passed `5/5` with no skipped case. The test family rebuilt every bounded fixture and restored Local/Dummy configuration plus fake-MTP module lifetime through ordinary cleanup.

## Five-process results

| Sample | Local Copy total / first mutation (us) | Local Delete total / first mutation (us) | Dummy total / first mutation (us) | Fake MTP total / first mutation (us) | Fake MTP cancel (us) | C isolated / D isolated / concurrent (us) | Concurrent / slower isolated |
|---:|---:|---:|---:|---:|---:|---:|---:|
| 01 | 172000 / 4577 | 78000 / 75425 | 31000 / 41356 | 625000 / 164529 | 62000 | 1187000 / 1172000 / 1171000 | 0.987 |
| 02 | 156000 / 4344 | 235000 / 232376 | 203000 / 29652 | 609000 / 162196 | 32000 | 1125000 / 1109000 / 1141000 | 1.014 |
| 03 | 172000 / 4057 | 78000 / 83157 | 31000 / 35311 | 703000 / 158093 | 62000 | 1187000 / 1172000 / 1234000 | 1.040 |
| 04 | 172000 / 5106 | 156000 / 155640 | 31000 / 36782 | 640000 / 162586 | 32000 | 1141000 / 1109000 / 1187000 | 1.040 |
| 05 | 157000 / 4648 | 94000 / 87224 | 62000 / 27255 | 671000 / 165329 | 46000 | 1203000 / 1125000 / 1172000 | 0.974 |

Nearest-rank p95 for five samples is the maximum. Local Copy total/first-mutation p95 was `172000/5106 us`; Local Delete `235000/232376 us`; Dummy `203000/41356 us`; fake MTP `703000/165329 us`; fake-MTP cancellation `62000 us`; and independent-volume concurrent wall time `1234000 us`. Every first mutation was below 1 second and occurred before discovery closure in each eligible fixture. Every task completed below 30 seconds with exact source/destination byte and item checks.

## Resource and policy gates

- Maximum ready queue: `128 <= 256`; discovery starvation: `0`; traversal-limit hits: `0`.
- Maximum discovery callback work: `131 us`; maximum discovery lock wait: `41 us`.
- Maximum bridge admission queue / retained entries / queued path bytes / metadata bytes: `0 / 0 / 0 / 3994`, below `256 / 4096 / 16 MiB / 8 MiB`.
- Fake MTP maximum concurrent backend calls was exactly `1` in every process. Cancellation completed in at most `62000 us <= 500000 us` while discovery and mutation were both live.
- C and D were independently marker-authorized. Both tasks made overlapping progress in every process. Worst concurrent/slower-isolated ratio was `1.040 <= 1.35`.
- `FileOps_ProviderCapabilityMatrix` passed `3/3` in `7578 ms` under Debug run `r4-a19-control-fileops_providercapabilitymatrix-20260902`, including Local UNC classification as the distinct uncontained `local-win32-smb` route and rejection before share access or mutation admission.
- `Causeway_BridgeSchedulingAndResourceContracts` and `Phase11_CrossFileSystemBridge` passed `3/3` in `6547 ms` and `12437 ms` under runs `r4-a19-control-causeway_bridgeschedulingandresourcecontracts-20260902` and `r4-a19-control-phase11_crossfilesystembridge-20260902`.
- The existing Phase 5/Fairstream Skip and just-in-time cases remain the behavioral controls for the one-way switch and bounded checkpoint release. R4-A19 produced no scheduler candidate, so it does not fabricate a second timing side or claim a relative throughput improvement; the earlier observation-only OR1 before/after comparison owns the only changed-code performance pair.

The pre-correction process `r4-a19-parent-release-20260902-05` exposed that fake MTP may surface a requested cancellation as `HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED)`, an established aborted-operation status. Commit `22892314` widened only that test assertion while retaining the 500-ms, serialization, closure, and resource gates. The five archived processes above were all rerun from the rebuilt receipt after that correction.

Result: **PASS — no scheduler mutation.** The current discovery reservation, low-water, half-rate, and checkpoint policy met the activated forward-progress, bounded-memory, provider-serialization, cancellation, contained-route, and independent-device gates.
