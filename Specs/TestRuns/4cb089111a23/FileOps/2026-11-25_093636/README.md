# I66 exact Leftover Cleanup consent

The x64 Debug `FileOps_ReviewRemediationContracts` selection passed setup, the case and
cleanup without desktop activation. Exact head `565368ea3` built with zero warnings and
errors. Run `20261125T093521Z-42164-9b645cb4e45b42389ea5961b3a2d81c8` used artifact
receipt `ad35f2ba19f376464bee83c27cc9c77cfd67b31aca436beca6529c8b01452ccc`.

The test exercised the ordinary Local ingress with two checked leftover files. The one card
explained permanent removal and the possible leftover. Cancel left both files intact.
Acceptance removed the two checked files after exact revalidation. The negative control
replaced one checked name before the worker acquired its pin. Both item results recorded
`ERROR_REVISION_MISMATCH` and no mutation. The original, replacement and unaffected
sibling files all survived. No modal or deferred card appeared for the changed selection.

| Original trace row | Event | Checked entries | Duration (µs) | Status |
|---:|---|---:|---:|---|
| 3436 | One combined consent accepted | 2 | 0 | `S_OK` |
| 3467 | Accepted selection revalidated before mutation | 2 | 429 | `S_OK` |
| 3690 | Changed object refused before consent | 2 | 0 | `0x8007051A` |

Previously, an exact permanent cleanup asked about the leftover and permanent removal
separately. The candidate verifies the reduction in questions and records one guard-cost
sample. Repeat timings for larger selections remain open.

The original `artifacts/selftest/last_run/perf/perf_metrics.jsonl` contains 46,648 rows and
16,829,748 bytes. Its SHA-256 is
`755d7896ff5943b1a03167840530c035849ef943e8ecd1bea8056efd46b92f45`.

The first candidate run `20261125T085948Z-85664-13c19a168d224b92a93c713de1dd596b`
refused an unchanged selection whose scope order differed from its receipt.
The second `20261125T091838Z-78548-5c9aa161bc7f489a86979cd5afe7c67b` reached the
replacement refusal and failed at a task-level assertion. The exact causes are now
verified on each failed item.
