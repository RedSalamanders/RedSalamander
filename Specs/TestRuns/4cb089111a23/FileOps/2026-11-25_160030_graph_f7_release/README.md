# I66 integrated x64 Release Graph and deferred-exit checks

The test-enabled full x64 Release build of `d6946334f` emitted zero warnings
and errors. This source tree matches [PR #285](https://github.com/RedSalamanders/RedSalamander/pull/285)
after its squash merge at `588a1814a`. Build receipt
`cd0a3c103bc7d37df9d4d87dd5df5ec215eff3b34a5a5c8180b8f8af718c355d` was
byte-verified by both runs.

| Non-activating selection | Governed run | Passed | Failed | Skipped |
| --- | --- | ---: | ---: | ---: |
| `R0fGraph_FakeGraphReadWriteCreateRenameRecycle` | `20261125T145915Z-14776-f1ca94b059c64e50a4d230cb6a1f9801` | 3 | 0 | 0 |
| `FileOps_DeferredExitRefusesNewWork` | `20261125T150030Z-78508-8a52eb5b23cd48d19009e8ab6f799dba` | 3 | 0 | 0 |

The first run used the owned Graph loopback fixture. The second included
the delayed Dummy provider's F7 prompt and worker, plus the Pack and Unpack
refusals after accepted exit. Both selections passed setup, the selected case
and configuration cleanup. The runner reported repository coverage
`NOT_EVALUATED`. Hosted x64 and ARM64 Release remain open.
