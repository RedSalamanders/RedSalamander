# Riptide R1-2/R1-3 Microsoft Drive Retry + Merge Partial

Date: 2026-06-15

Scope:
- R1-2: Microsoft Drive Graph mutating request retry behavior.
- R1-3: Microsoft Drive directory-merge child failure classification.

Fix summary:
- Factored Graph throttle retry status/delay handling into shared helpers.
- Enabled the existing Retry-After loop for `PATCH` move/rename and `DELETE` item mutations.
- Kept `POST` directory create out of blind transport retry; on 429/503/504 it re-queries the target child first and treats an already-created folder as success before retrying.
- Changed merge type-mismatch overwrite handling from hard abort to skipped child + `ERROR_PARTIAL_COPY`, allowing unrelated siblings to continue moving.
- Added a Debug-only Microsoft Drive fake Graph transport and plugin selftest export exercised by `PluginContractTests`.

Evidence:
- Build: `msbuild-FileSystemMicrosoftDrive-20260615_183156_841.log` passed with 0 warnings / 0 errors.
- Build: `msbuild-PluginContractTests-20260615_183210_657.log` passed with 0 warnings / 0 errors.
- `GREEN_PluginContractTests_debug_graph_selftests/PluginContractTests.txt`: `PluginContractTests passed`; Microsoft Drive debug selftests reported `passed=25, failed=0`.
- `GREEN_FileOps_ProviderCapabilityMatrix/`: `FileOps_ProviderCapabilityMatrix` passed `3 passed / 0 failed / 0 skipped`.

Covered scenarios:
- Child `PATCH` move receives an initial 429 with `Retry-After: 0`, retries through the production retry loop, and completes the folder merge.
- `DELETE` receives an initial 429 with `Retry-After: 0`, retries through the production retry loop, and removes the item.
- `POST` directory create receives a 503 after the fake server has already committed the folder; provider re-queries the target and returns success without replaying the create.
- Folder merge with a folder-vs-file type mismatch answered `Overwrite` returns `ERROR_PARTIAL_COPY`, preserves the unresolvable source child, leaves the destination child intact, and still moves a non-conflicting sibling.
