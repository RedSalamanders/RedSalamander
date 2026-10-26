# Waiting cards and sorting merge preparation - 2026-10-20

The user requested merging all current changes, including the waiting-card fix,
queued destination-change removal, Debug row reordering, sort-paint measurement,
tooling, specifications and their retained evidence.

## Integration checks

- The five focused tooling files passed 306 tests, with zero failures or skips:
  ToolInventory, ToolingGovernance, SpecInformationArchitecture, ShowPerfRuns and
  TestHarnessSourceContracts. `pester-results.json` retains the original result.
- Tool inventory: 197 discovered and classified files, zero blocking findings.
- Specification inventory: zero blocking findings after adding exact-file admission
  for the waiting-card completed record. Its new guard rejects a date-shifted sibling
  and an unreviewed companion. Owning documentation and inventory were updated.
- All 495 newly staged original evidence files passed the changed-archive contract.
- All 27 changed native source files pass the CI formatter. Fourteen required
  formatting. Every formatter replacement was checked against original UTF-8 byte
  ranges before applying it and was whitespace-only; normalized byte readback and
  formatter idempotence also passed. `format-proof.json` records before/after hashes
  and replacement counts. Pre-format sources remain in the marked sandbox at
  `Z:/RedSalamander.Perf/runs/merge-all-20261020/scratch/Merge/Preparation/before-format`.

## Qualification boundaries

The original waiting-card and sorting archives retain their original source and
binary identities, failed attempts, focused native outcomes and performance limits.
Those byte-identity statements describe their recorded pre-format snapshots. The
merge preparation did not run new native binaries after formatting and does not
promote those receipts to the final tree or a repository-wide pass. Whitespace
normalization changes source bytes without adding product behavior changes.

Affected planning was inspected. Full, focus-taking, ASan and ARM64 qualification
remain pending where recorded by the original plans. No idle-host, speedup or
performance-budget claim is added by this integration record. CI after the push
is separate pending validation.
