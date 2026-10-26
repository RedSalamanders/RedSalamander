# Existing destination Move failure text

Date: 2026-10-16. Machine: 4cb089111a23. Profile: native x64 Debug.
Branch: ej/move-existing-file-message. Base: a48e7d7880cb37ea6898d462ae188a6213a5f3f8 plus the saved source diff.

## Change and protected scenario

A Managed Move can stop before mutation when its existing destination cannot expose object
identity (for example a NAS without FileIdInfo). The old Last note was the generic Windows
"The request is not supported." The new typed problem says the destination already exists,
explains that replacement cannot be validated safely, and states that both objects were kept.
The original HRESULT stays unchanged. Only this failure path performs an additional existence
read; this diagnostic never grants overwrite or destructive cleanup authority.

The canonical File Operations and popup specifications own the enduring contract.
`source-diff.txt` records the implementation, five-language resources, test and spec changes.
Two pre-existing formatting differences in the touched FileOperations.cpp were normalized.

## Validation

- `candidate-final/`: `Test-Changes.ps1 -Suite FileOps -CaseFilter FileOps_ReviewRemediationContracts`
  passed setup, the regression phase and cleanup (3 passed, 0 failed, 0 skipped). Build passed
  with 0 warnings/errors. Receipt: cdb0bffaf9cb225bb68b43ebff3fbb68a27e00e8b208fdcd7d79a07b51d0d702.
  The actual guard/executor/finalizer/completed-summary seam covers different bytes, identical
  bytes, a whole destination root without identity support, an absent destination and unsupported
  source identity. Every source/destination byte stream is preserved, no item mutation callback
  runs, and no source is removed. The specialized diagnostic is emitted exactly three times.
  The running note agrees with completion; completed note and Issues follow French/English live
  language changes and survive a trailing aggregate task.result diagnostic.
- `candidate-initial/` retains the earlier successful candidate receipt and trace, before final
  formatting. It is not promoted as evidence for a different artifact identity.
- `pester-results.json`: ResourceLocalizationContracts and TestHarnessSourceContracts passed
  214/214 with no skips. Pester TEMP/TMP were redirected into the existing marked test sandbox.
- Pinned clang-format dry-run and git diff --check passed on the final sources.
- Whole TestRuns archive inventory passed for 13,844 files. This curated archive also passes
  explicit pre-stage archive validation.

## Broader conflict coverage failed

`normal-conflicts/` preserves the FAILED canonical run:
`Tools/Run-AllTests.ps1 -Suite FileOps -CaseFilter FileOpsFamily_Phase09_ConflictPrompt`.
Run: 20261016T213247Z-40136-8f0189f6619c4612b378f9a7ab86b0ce.
Its build passed with 0 warnings/errors, receipt
495031729cbaa82f311956d986516f97b200e413acf9f5e898bbf521d48f8673.

Overwrite/replace-readonly and Apply-to-all passed. The KeepBothNestedCacheEligibility phase
failed in its copy-parallel scenario because the nested collision was reached before its
"earlier" sibling had completed. Six remaining phases were not reached; their individual
skips and reasons remain in fileops-results.json. The aggregate is 4 passed, 1 failed, 6 skipped
(including setup/cleanup). This is not an overall green conflict suite or Full repository result.

The failing assertion is in the existing parallel COPY scenario; the new diagnostic branch
requires MOVE. The retained raw metrics contain zero DestinationIdentityUnavailable rows and
zero ERROR_NOT_SUPPORTED guard rows in that run. This locates the observed failure outside the
new diagnostic path; it does not establish whether the failure predates this patch. Its root
cause and the six skipped phases remain unqualified.

A prior -SkipBuild attempt was blocked because the available receipt was targeted rather than
full-solution. A Test-Changes attempt deferred this desktop-serialized family without launching
it. Neither is counted as a runtime pass. The final canonical run used --selftest-no-activate.

## Performance evidence and limits

The compact candidate metric streams retain every matching row, in original order, for
fileops.identity.guard_us, FileOps.Conflict.DestinationIdentityUnavailable, and
FileOps.SelfTest.UnverifiableDestinationCollisionMs. Each metric-selection.json binds the raw
stream by SHA-256, byte count and row count. Full raw streams were copied before runner cleanup
into Z:\RedSalamander.Perf\evidence\move-existing-message\candidate-initial and candidate-final;
the broader failure's full stream is retained in the sibling normal-conflicts directory.

The final five elapsed GetTickCount64 samples are 0, 0, 16, 0 and 16 ms. These are observational,
coarse clock samples, with no speedup or throughput claim. See baseline-observed.md for the
observed pre-fix failure and the explicit loss of that run's raw archive to automatic stale-run
cleanup; no quantitative baseline comparison or performance qualification is claimed.

The NAS failure is reproduced by deterministic Local identity-unsupported injection. The live
user share was not touched or qualified. Release, ASan, ARM64 and Full/PR gates were not run.
The runner also reported existing sandbox disk-audit findings (32-33); they were not altered.
