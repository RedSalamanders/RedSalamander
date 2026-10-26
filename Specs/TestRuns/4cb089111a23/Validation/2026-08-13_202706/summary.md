# Operation Startrail validation-planning evidence

Same-machine, same-fixture measurements on `4cb089111a23`, x64 Debug, with 59
changed workspace files and the 18-entry Full plan. These are median/maximum
claims, not p95 claims; the archive therefore does not invoke the separate
200-sample percentile rule.

- Warm impact + schema planning: median 614 ms, maximum 773 ms (limit: 2 s).
- Cold snapshot + impact + schema planning: median 3,482 ms, maximum 3,641 ms
  (limit: 10 s).
- Full-plan entry fingerprints: median 32,317 ms before receipt-binding reuse;
  median 279 ms after, plus one 1,749 ms receipt-closure binding per run.
- Evidence-history lookup: median 1 ms across 21 samples.

The requested FileOperations-only Commands avoidance is explicitly blocked for
Full evidence. Commands family selection now exists and its 14 native family
inventories form an exact duplicate-free union of the 845-case broad inventory,
but every family still executes from the monolithic `RedSalamander.exe` runtime
closure. Consequently family entries are `Never` cacheable and repository
status remains `NOT_EVALUATED`; the broad Commands process is not reused from a
logical family result.
