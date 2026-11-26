# I66 R0e Batch Rename fixture review

The Batch Rename plan fixture in PR #312 now supplies both destination paths and collision keys, the captured parent, and one schedule layer. The complete plan rejects two bindings to the same physical source with `E_INVALIDARG`. The same plan accepts two distinct physical sources with `S_OK`. The witness runs before the R0e timer. Its exact normalized source at `20333bca6eeb976ac2bf03938d4d369b497fb32a` matches the x64 Release witness run's `0ff0b5c4bea0c01b40cb6cae930a71b225348fa4`.

| Normalized source identity | Previously reviewed fixture | Revised fixture |
| --- | --- | --- |
| Whole case | `2E5558B523D99D73528B2664EA56CE2E1A7733952CB572CF62FF00E2636FA329` | `AC19290438FF0A6C57A810263AD6C371DC34C4531FF45FBD12147A6EB394227F` |
| Prefix through the predicate | `B5580DB8FC0D352F49ACAF2FE5E99100816D7CFB8B20AC676A2F06A2A93AD488` | `E5BDBF77A737E15012F0FD30DA7741B07739ECD83393B62623CD3FB9B59BBA59` |
| Diagnostic | `E593F40EFA4550E1227389139A42D84A4ACA8CD8CF152886F7B0EBE2947CA9A0` | Same bytes |
| Remaining assertions | `681389C78B53E92548FD3D86563E6AAECC0AB5B3AE00E21AFAC8C6D0FC46FA01` | Same bytes |

The measured code and predicate from `kR0ePerfIterations` through the end of the predicate match byte for byte. The source validator owns candidate eligibility and keeps previously reviewed identities. The hosted command compares both fixture projections before building the baseline.

| Check | Result |
| --- | --- |
| Exact x64 Release provider matrix, `20261126T085554Z-59684-4b5dfc4f169f4b38b624ce3bf9eed32e` | Three passed. Route containment took 327,823 µs with 4,096 parses, 4,096 decisions, one rejection and zero provider calls. |
| Fresh x64 Release PrePush, `20261126T091500Z-67596-ef842a55b4f247ec8f13e1973ed3a1ae` | 1,799 passed, two failed, 29 skipped and four interactive entries pending. Both failures rejected the unregistered case before archive validation. |
| First focused correction, `20261126T102247Z-71884-6c71e5a7d87a46eca97100aeac2a442a` | 19 passed, one failed. The standalone identity passed. The hosted command's separate candidate list still refused the case. |
| Corrected focused Pester, `20261126T102400Z-47644-5cf79fb848d64d1daba52f58efefb7f9` | 20 passed, zero failures or skips. The hosted command reached the injected archive failure and retained the run ID and log. Source controls refused changes to the Rename fixture, workload, ceiling, diagnostic and activation guard. |

The focused suite's raw summary is under the completed run's `artifacts/selftest/last_run/r0e-pester-summary.json`. The tested module SHA-256 is `13741FB93BD8542D6E865FACA8CC472C5A8CFF63680B93FC8A3CE2DA49958CE5`. The tested hosted command SHA-256 is `B2DC414E3AAFFE0FC35498C2AB8FF08B12B63F3077C2D5CFA89B16740FD8B071`. The exact native run is also recorded with the [worker interlock witnesses](../2026-11-26_095011_worker_interlock_witnesses/README.md). The five-process hosted diagnostic, the interactive entries and both native Release qualifications remain separate gates.
