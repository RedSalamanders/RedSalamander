# I66 repeat-decision checkbox

The same-machine x64 Debug pair ran `FileOps_RepeatDecisionContracts` with no desktop activation.
The baseline and candidate each passed setup, the decision case and cleanup.
The candidate app built with zero warnings and zero errors and used receipt
`cb918644a8d7b2cbcb0301bad0a73eb1447eb73b8a9909b3ceb7da77cbec65ef`.
Each sample evaluates 16 passes over 26 conflict classes and 64 permission masks. The five
samples follow one full warm-up. The change in offered actions alters the expected matrix checksum.

| Variant | Run | Measured batches (µs, in order) | Median (µs) |
|---|---|---|---:|
| Baseline | `20261124T184052Z-90464-d6e58d0eebdb48b3ba204f5228c24b8d` | 44889, 43617, 43915, 44842, 44655 | 44655 |
| Candidate | `20261124T184149Z-90464-a9deeb3762ba496f8d0293654da68d4b` | 41740, 41665, 43133, 43032, 47443 | 43032 |

An earlier implementation scanned the repeatable actions during each policy construction and
raised the median from about 42 ms to 82 ms. The candidate derives the skip-only label when a
question publishes. The final synthetic matrix median was 3.6% lower than the paired baseline.

The same candidate passed the no-activation mirror UI Automation case
`20261124T184247Z-90464-06ff3b60510b40da9b2f9abb4172a897`. It exercised the real toggle and its
checked state. The previous implementation passed the broader remediation case
`20261124T182349Z-103804-54f36d7b122d4a02b31d42eff6e176d3`. The later remediation run
`20261124T184343Z-90464-6cdb486467ea431dbc2bb66d241cc7a2` failed when sharing error 32
prevented reading a Recycle Bin metadata record during per-item cleanup. The failed result,
trace and exact owned fixture path remain in the run evidence. Full interactive coverage remains open.

Each metric file retains original lines 249 through 254 in their original order. These are the
repeat-scope control and five batch timings. The complete metric files remain in the unified
sandbox run archives.

| Variant | Original rows | Original bytes | Original SHA-256 |
|---|---:|---:|---|
| Baseline | 1718 | 638950 | `d9796fea8e59de3bd37854d817cfd6a902e95ad56e8518a5ae0492de4e651170` |
| Candidate | 1701 | 616770 | `8b4a52e64064dada19dee6763fcea659fcf4fe58d58af5bd2e30cf2bd2ef38c8` |
