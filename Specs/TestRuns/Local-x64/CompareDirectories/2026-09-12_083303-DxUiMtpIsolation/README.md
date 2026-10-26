# MTP journal isolation regression

The old executable fails `mtp_property_fetch_is_batched` when its synthetic journal
already occupies all 64 quarantine names (ERROR_ALREADY_EXISTS). It passes with fresh
scratch state. Both baseline runs predate the isolation change; their binary hashes,
results and traces are retained here.

At b97bbc1a, the corrected Debug executable passes with the same saturated input.
All 65 caller journal files retain their hashes. The full synthetic MTP family then
passes two repetitions in shuffled order (seed 190912): **118 passed, zero failed,
two skipped**. The two skips are the explicitly unconfigured physical-device smoke.
Each attempt receives a fresh governed journal root and restores the previous root.

The full Debug solution build reports zero warnings and errors. These focused results
establish the fixture correction; they do not replace Fresh Full or native ARM64
qualification. Final rebase and qualification remain open while upstream CI compilation
fixes are in progress. No user journal files were deleted or modified.
