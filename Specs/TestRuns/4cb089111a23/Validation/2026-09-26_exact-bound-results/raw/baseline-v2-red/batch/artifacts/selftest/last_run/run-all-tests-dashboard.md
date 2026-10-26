# Run-AllTests Case Dashboard

- Run: `20260923T204313Z-68856-6c737d5a44cd4772a2cdffab03c578a0`
- Suite: `FileOps`

## Top Slowest Cases

| Harness | Case | Source | Status | Classification | Duration ms | Seed | Attempt |
|---|---|---|---|---|---:|---|---:|
| FileOperations | Setup | case | passed | UNCLASSIFIED_FAILURE | 1187 |  | 1 |
| FileOperations | FileOps_BatchRenameExactResultCarrier | case | failed | UNCLASSIFIED_FAILURE | 204 |  | 1 |
| FileOperations | Cleanup_RestorePluginConfig | case | passed | UNCLASSIFIED_FAILURE | 31 |  | 1 |

## Failure And Retry Evidence

| Harness | Case | Source | Status | Classification | Seed | Reason |
|---|---|---|---|---|---|---|
| FileOperations | FileOps_BatchRenameExactResultCarrier | case | failed | UNCLASSIFIED_FAILURE |  | Exact batch-rename result lacked the accepted prompt reference or truthful outcome axes. |
