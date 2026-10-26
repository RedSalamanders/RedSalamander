# Nested archive Copy regression evidence

Candidate: master base 5141305d9c1d284c031d325d2e448fcd4a13a1d9 plus the working-tree fix and tests, x64 Debug, machine 4cb089111a23, 2026-10-16.

The archive provider now advertises both slash spellings that its reader already accepts. Preparing consequently derives the selected leaf correctly and retains the destination provider's existing validation.

Command: `./Test-Changes.ps1 -Suite FileOps -CaseFilter FileOps_ArchiveNestedCopy -Interactive -SkipBuild -Configuration Debug -TestRoot E:\RedSalamander.Perf`.

Run: 20261016T142300Z-90968-2a4d7471cccf4ae09d5e6c57bb2c5a06. The runner verified solution receipt 0b6c2c03982146f57ead6ec08d8819c634fa4b410887f39bea420ba35e1fdbc3. Both new cases, Setup and Cleanup pass: 4 passed, 0 failed, 0 skipped. The full Debug solution build has zero warnings/errors. The changed Release plugin also built with zero warnings/errors; its separate receipt is retained. This is not complete Release runtime qualification.

The unchanged raw perf log contains 24 successful `FileOps.SelfTest.ArchiveNestedCopy` rows: 12 local and 12 over the guarded loopback SMB alias. Each target covers ZIP, 7z and four-volume RAR with a control path, nested canonical path, exact Windows FolderView join and all-backslash nested path. Assertions require exact output bytes, a single selected filename, no extra directory or stage, and continued pre-mutation rejection of reserved/trailing-dot names and invalid child names. Both targets copy 28,427 bytes in total.

| Target | Copies | Minimum | Median | Maximum |
|---|---:|---:|---:|---:|
| Local | 12 | 22.716 ms | 27.7135 ms | 41.932 ms |
| Loopback UNC | 12 | 28.176 ms | 33.8755 ms | 59.544 ms |

These are small correctness-fixture timings from one candidate run, including host bridge and observation overhead. They do not establish a throughput improvement, NAS performance, or integrity of the user's large movie. No comparable successful pre-fix Copy timing exists because the path contract rejects the affected source.

Standalone contract negative control and candidate: `../../../Local-x64/FileOps/2026-10-16_ArchiveNestedCopyContracts/` preserves four new assertion failures against the old on-disk Debug plugin and the candidate's full PluginContractTests pass, including all seven new assertions. That candidate used the same production fix and contract-test source; the receipt there binds its earlier complete test source state. It is not reused as full qualification for the final test executable.

The source-contract Pester file passes all 201 checks after the final fixture change; its unchanged NUnit XML and summary are included. Earlier focused attempts selected an unsupported split-7z outer container and failed to locate the requested nested entry. Their original failures remain in `../../../Local-x64/FileOps/2026-10-16_ArchiveNestedCopy_FirstFailure/` and `../../../Local-x64/FileOps/2026-10-16_ArchiveNestedCopy_SplitFixtureFailure/`. The final matrix uses a real multipart RAR with pinned provenance and hashes in `Plugins/FileSystem7z/Tests/MultipartRar/README.md`.

The original runner summary reports two disk-audit notices for manually named evidence directories beneath the marked sandbox. Both were relocated, without deletion, into this run's artifacts/auxiliary directory. The subsequent canonical disk audit is clean (0 issues), preserved separately as disk-audit-after-relocation.json. The original summary is unchanged.

Broader affected validation was interrupted after its tooling profile reported a source-contract naming failure (937 passed, 1 failed). The guard required the result variable `uncStatus`; that name is repaired, and the owning source-contract file now passes 201/201. Partial Compare execution had no failures but did not complete. No Full repository pass, complete affected pass, ARM64/ASan qualification, or NAS Copy is claimed.

