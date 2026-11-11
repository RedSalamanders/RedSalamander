- When a Move answered `Delete source` meets a source folder without a Recycle Bin, the task now asks once
  whether to delete such sources permanently (`Permanently delete`, `Keep source`, `Cancel`) instead of
  always keeping them. The answer covers the rest of the task.
  - The question is the Move's Recycle-refusal consent card: Cancel is the default, no Apply to all, no
    `More...`. One worker asks; the answer is kept on the conflict arbiter (`sourceNoRecycleBinAnswer`).
  - `Permanently delete` binds the source before the question and again with delete access after it,
    requires the same identity, size and last-write time, repeats the distinct-file proof and consumes the
    bound object with `DeleteIfUnchanged`. A changed source, a second name of the destination, a read-only
    source or a refused delete keeps the source as a skipped item with a Warning; an unknown outcome is an
    Error.
  - `Keep source` now records an Information note instead of a Warning, so the user's own choice does not
    count as a problem.
  - Five strings in five resource files; the Local provider's test seam
    `REDSALAMANDER_FILEOPS_RECYCLE_UNAVAILABLE_DIRECTORY`; new self-test
    `RunFileOperationDiscardSourceWithoutRecycleBinSelfTest` in `FileOps_ReviewRemediationContracts`.
  - Not covered: a source whose location cannot identify files (Delete source is not offered there) and
    sources on other providers. Not built or run when written.
