# C0 native mutation receipts — focused verified candidate

Run `c0-native-receipt-matrix-verified-20260905`, Debug, `C0_` prefix:
5/5 passed, no skips (three behavior cases plus setup/cleanup). Build receipt
`667f6840edaeddcc00bec93e28c2cd6550f3728e35a119b60724da9144104c44`;
zero warnings/errors. Base `58f22d0e` plus working changes is bound by the runner.

| Behavior case | Verified facts |
|---|---|
| `C0_GDriveCommittedCopyResponseLost` | One backend request/commit; both files exist; no conflict prompt; uncertain publication, retained source, localized uncertain summary. |
| `C0_GDriveCommittedDeleteResponseLost` | One DELETE request/removal; source absent on the backend; no conflict prompt; indeterminate host result with unknown source disposition. |
| `C0_NativeCopyFailurePreservesKnownAxes` | Unknown cleanup retains known non-publication and original destination bytes. Post-publication attributes failure retains Published and replacement bytes. Source survives both. Each task has only its ordinary overwrite prompt. |

The prior `2026-09-05_c0_native_receipts_fault_location_red_debug` archive documents
the misplaced pre-publication attributes hook and failed first matrix. The corrected
hook now fires after actual publication/backup cleanup; the existing direct-provider
overwrite test also asserts replacement bytes. All raw results, trace, metrics and
runner aggregate are retained here. Request/prompt counts are correctness and
amplification guards, not throughput or percentile evidence.

This is not plan closeout. The complete conflict family, publication and direct
Local overwrite regressions, provider/walker matrix, Release, remaining C1–C5 work,
human checks and final Fresh Full remain required.
