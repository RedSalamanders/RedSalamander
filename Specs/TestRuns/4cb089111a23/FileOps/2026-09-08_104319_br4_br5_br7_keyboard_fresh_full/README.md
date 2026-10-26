# BR-4 / BR-5 / BR-7 / Keyboard Fresh Full closeout

Source: `892f92358e9a4bab17fa07ab81d0604e36849885` plus the BR-4/BR-5/BR-7/BR-8 worktree.
Governed run: `20260908T104319Z-97448-ff25c4b1b0d545fe9b3654fc1edfb85f`.
Command: `Tools/Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -AllowAlternateVolumeTestRoot -TestRoot Z:\RedSalamander.Perf`.
Profile: test-enabled x64 Debug; initial build and required re-attestation both reported zero warnings/errors.
Accepted build receipt: `4cefc3636a4084de4af21b81abab35941efa94ebfcf7ab9daf1d19f11ca32d45`.
Workspace snapshot: `586e01641d3b534ad321b5b3a4be717601c6398c4e08fde66d3fd85223a24e10`.

## Verdict

Repository and change verdicts are **PASSED**. All 35 entries executed and were
promoted: zero reused, provisional, invalidated or incomplete entries. Commands
has a fresh broad seal. The runner recorded **2,113 passes, zero failures, 52
declared skips**, 2,165 total, in 4,012,739 ms. No quarantine repair or failure
reclassification was needed. This is one Fresh Full invocation, not composite
or isolated-repeat clearance.

| Suite | Passed | Failed | Skipped |
|---|---:|---:|---:|
| Commands | 896 | 0 | 2 |
| FileOperations | 180 | 0 | 20 |
| CompareDirectories | 242 | 0 | 30 |

Every case status and skip reason is retained. Commands skips its two explicit
ViewerSpace opt-ins; FileOperations retains 18 unavailable remote-profile checks
and two 7z cases owned by Compare; Compare retains its 30 prerequisite skips.
The runner also reports 22 pre-existing test-root layout audit residuals; these
are preserved in the aggregate and were not deleted or waived by this work.

## Required witnesses

| Case | Result and duration |
|---|---|
| Original Keyboard reordered/resized copy after search round trip | Passed, 686 ms |
| Keyboard displaced-native-focus control | Passed, 743 ms case; exact input reacquired in 125,000 us |
| Original navigation popup edit route | Passed, 284 ms |
| Popup reentrant-close control | Passed, 340 ms case; folder focus returned in 15,000 us, no premature restore consumed |
| Popup owned-window activation | Passed, 300 ms |
| `Beeline_CopySkipLinksKeepsPlaceholders` | Passed, 281 ms; D: NTFS, actual Cloud tag `0x9000601A`, attributes `0x420`, all 65,536 bytes copied through one worker, junction skipped, no hydration prompt |
| `BR5_CurlHostDirectoryMove` | Passed, 140 ms case; 78,000 us host operation, exactly one RNFR/RNTO, nested/empty/sibling state preserved |
| `BR5_CurlHostDirectoryMoveRefused` | Passed, 375 ms case; 172,000 us host operation, one RNFR/RNTO, source intact, no relay/delete/retry; host result remains conservatively Indeterminate |
| `R4A19_DiscoveryIndependentVolumes` | Passed, 3,875 ms, authorized Z: ReFS and D: NTFS roots |
| `BR3_CancelQueuePausedTransfer` | Passed, 9,343 ms fixture |
| `Phase7_ParallelCopyMoveKnobs` / `Phase7_SharedPerItemScheduler` | Passed, 38,515 / 1,062 ms |
| Tasks-quiet / popup-host shutdown | Passed, 1,031 / 593 ms |

The fully hydrated placeholder needs neither a cloud account nor a network
provider. The fixture restores the previous process/thread placeholder modes
and unregisters its temporary sync root before cleanup. The production repair
routes a regular-file reparse into the existing bounded worker queue; the
parallel producer no longer attempts a synchronous copy after releasing its
primary buffers.

## Evidence preservation

Case results are unchanged copies of the runner's immutable entry results.
Commands and FileOperations traces come from their suite-specific automatic
archives because the live `last_run` directory rotates between suites. Aggregate
v2/v1, case history, plan, decision, terminal run state and final workspace snapshot
retain the accepted receipt and promotion identities.

`perf-compaction.json` binds each original suite stream's byte/row counts and
SHA-256 to its selected output. All `FileOps.SelfTest.*` rows and the two exact UI
focus metrics retain original bytes and order within their own suite streams.
No general throughput improvement is claimed. Git attributes preserve these
curated bytes on checkout. Subsequent repository edits record this closeout,
repair plan links, preserve archive bytes and admit the exact new Done path in
the protected-history allowlist. Application/plugin source stayed fixed.

The post-gate closeout metadata has separate focused qualification:
`SpecInformationArchitecture.Tests.ps1`, `DocumentationDriftContracts.Tests.ps1`,
`ToolInventory.Tests.ps1` and `ToolingGovernance.Tests.ps1` passed **75/0/0**.
The new admission test accepts this exact Done record and rejects an unreviewed
sibling. `post-gate-closeout/` preserves every focused case result and the summary.
These counts are separate from the Fresh Full result above. Spec and tool
inventories report zero blocking findings.
