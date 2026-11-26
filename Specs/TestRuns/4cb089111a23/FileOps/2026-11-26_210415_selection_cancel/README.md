# Large Local selection counting cancellation

The [hosted x64 Release Full diagnostic](https://github.com/RedSalamanders/RedSalamander/actions/runs/45841032537)
passed the six Local selection concurrency scenarios. Its 4,097-file counting check verified
the discovered totals, duplicate report, loss of byte certainty after a changed report,
unchanged source and stage cleanup. The check returned `ERROR_PARTIAL_COPY` when the
fixture expected `ERROR_CANCELLED`. The fixture had been sampling discovery closure while
the worker could already admit transfers.

The existing discovery publication hold now also covers the fully counted multi-file prepass.
The fixture observes the live totals, exercises both reporting controls and asks the real
task to cancel before it releases the worker. An empty destination verifies that no file or
stage published. The original single-leaf tests retain their existing hold. The prepass metric
is recorded before the hold.

| Source | Non-activating x64 Release run | Passed / failed / skipped | Discovery |
| --- | --- | --- | --- |
| `1611a9086` | `20261126T194807Z-61936-d68ea92ec05a41928647fac4dcfbbf45` | 3 / 0 / 0 | 94,000 µs |
| `8ac8ed49e` | `20261126T200021Z-22772-ddeeb8289ee2420f81d17f498b795f56` | 3 / 0 / 0 | 109,000 µs |
| `59c3d985165739f9d290b54d876ea40e2d0a0428` | `20261126T200415Z-59700-c9a1c9fba53b4f329e8a688104eb9ae6` | 3 / 0 / 0 | 78,000 µs |

The exact narrowed candidate built with zero warnings and errors. The runner verified receipt
`c4c1b176261074182bd7264e65df8cc7062537d7d328bcd5fca9fb6bb8ae1e32`.
The original six scenarios still recorded the expected four, four and one active files for
both folder and individual-file selections. The large selection reported zero bytes while
discovery was open, the expected `ERROR_CANCELLED`, unchanged source and an empty
destination. The native result SHA-256 is
`e000357e04f9bb270b498d53ef01be141c470e640ff56bbee392d32224096d8a`.

All runs used owned fixtures under the marked local sandbox. The earlier failing hosted
result remains archived with the [pointer evidence](../2026-11-26_203625_disclosure_start_now/README.md#hosted-x64-diagnostic).
Hosted native platform and full desktop gates remain open.
