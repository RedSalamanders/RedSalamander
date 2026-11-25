# I66 Pack and Unpack source retention

The owned ZIP roundtrip approved Pack's delete-after option for a selection
containing a file symbolic link. The ZIP contained the regular files and
omitted the link. The test confirmed that Pack kept the original files and
link, reported a partial result, and queued no cleanup Delete. The same case
also confirmed that Unpack kept the selected archive and the conflicting
destination when **Skip existing files** left requested entries unextracted.

- Full x64 Debug candidate `63504c595` built with zero warnings and errors on
  receipt `7f9773fdd3ea3fee82ebbd6fe822f77ec4b3535b02897e2d7bb18eeada565c94`.
- Non-activating archive case `20261125T161424Z-8584-f111c4d7cf174542b2d63fa66cce8229`
  passed once, with zero failures and skips. The native fixture retired its
  exact symbolic link and the governed evidence archive completed.
- Run `20261125T161033Z-53224-ef2fc0ef8062488782b3a0da3c372335` passed
  67 runner-plan checks and 276 source and tooling checks. The 12 resource
  checks passed in `20261125T155304Z-70332-87ef0be92b384d82801e4eaa371f7df6`.
  The resource inputs were unchanged by the teardown correction.

The first exact native run `20261125T160105Z-86300-d50009f30fab4a5db0e0f568a44de3aa`
passed the case. The run wrapper failed because the fixture left the
symbolic link in the archive tree. It was preserved as an interrupted
run after verifying and retiring the single owned link. The target digest
and native result were preserved. Candidate `63504c595` added scope-bound
teardown using the existing no-follow test helper and passed the clean rerun.

The exact PR head `e7a6b1f61` also built for x64 Release with zero warnings
and errors on receipt
`522109fc9f2e86e400bfc45be7fe54e7ada3e91305d1eb7599e3537d917d86fa`.
The non-activating selection passed in 845 ms with zero failures and skips
in `20261125T163138Z-85480-2dddd82cea854964a77d03ceb8eb31a2`. The run
archived normally. [PR #291](https://github.com/RedSalamanders/RedSalamander/pull/291)
merged at `64b7c09d2` after the required checks.

Repository coverage is `NOT_EVALUATED` for the selected runs. The plan tracks
worker migration and original-source proof.
