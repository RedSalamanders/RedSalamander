# Search progress diagnostic volume — Debug component evidence

The same public Repository.Query fixture, local_index_core_progress_counters_are_batched,
ran against the original and batched producers. Its source blob remained
a96d83fecca970a557db0d24f91fcea4620906a8. The fixture uses 4,096 invented files inside
the marked C:/RedSalamander.Perf sandbox and the existing traversal override.

| Successful-query observation | Original producer | Batched producer |
| --- | ---: | ---: |
| Appended metric bytes | 5,880,051 | 10,750 |
| Calls represented | 8,197 | 8,197 |
| Suppressed calls represented | 8,188 | 8,188 |
| Actual callback records | 9 | 9 |
| Forced callbacks | 5 | 5 |
| Call-count rows | 8,197 | 9 |
| Suppression-count rows | 8,188 | 4 |

The original producer at b1264efba9bf31998d0384f8e8b5fc6c4c8e8216 failed the new
diagnostic-volume/row-bound requirement as expected. The candidate at
20c564af6ee26bcd226758f6cbc183b7d0b04c5d passed with no skips. Both retained E_ABORT
for callback rejection, ERROR_CANCELLED for early cancellation, all 65 cancellation calls
and 62 suppressed calls, and silence for a null callback. The cancellation path retained
its terminal delta flush. The successful query returned the complete expected name set.

The copied case results and complete trace files preserve every declared case status and
reason, original bytes and line endings. evidence-index.json binds them and the external
raw JSONL streams to SHA-256 and byte counts. Raw metrics remain in the immutable local
run archives; this directory contains compact review evidence only.

Both x64 Debug solution builds completed without warnings or errors. Their accepted
receipts were ca5c0efefe20e5c4badf5db438e00e22b0f7b261a2e3dc3875d2103e4b649cfa
and e34a1e2983af67fb170a1f7bbe0e89ceb63dda857ed3c68c45fcfd3d15fd96ad.
The original-producer run is 20261024T162715Z-98188-c4ad6348d6384111a7557e3d018d3230;
the candidate run is 20261024T163728Z-24504-051f94a8ab7946608e5eb1d7f417f755.

This establishes lossless counter accounting and reduced diagnostic volume for this
Debug fixture. It is not Full repository qualification, a production search throughput
claim or paired Release performance qualification. Read-only specification/planning work
overlapped the baseline; elapsed times include metric inspection and are not used as a
causal product-speed comparison. The earlier failed merged-source Full remains failed.
