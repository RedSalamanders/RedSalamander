# Run-AllTests Case Dashboard

- Run: `20260923T204348Z-68856-62dfeaef8b5d401bb053601eaab18209`
- Suite: `FileOps`

## Top Slowest Cases

| Harness | Case | Source | Status | Classification | Duration ms | Seed | Attempt |
|---|---|---|---|---|---:|---|---:|
| FileOperations | Setup | case | passed | UNCLASSIFIED_FAILURE | 1187 |  | 1 |
| FileOperations | Phase10_DeferredConsentAndRecycleEscalation | case | failed | UNCLASSIFIED_FAILURE | 250 |  | 1 |
| FileOperations | Cleanup_RestorePluginConfig | case | passed | UNCLASSIFIED_FAILURE | 31 |  | 1 |

## Failure And Retry Evidence

| Harness | Case | Source | Status | Classification | Seed | Reason |
|---|---|---|---|---|---|---|
| FileOperations | Phase10_DeferredConsentAndRecycleEscalation | case | failed | UNCLASSIFIED_FAILURE |  | Exact Recycle escalation needs separate accepted-choice and bound-delete result diagnostics. |
