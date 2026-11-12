- A folder moved onto a folder of the same name at a location that cannot identify files (a share
  without file IDs) now merges, as on any other drive, instead of being refused.
  - The folder itself raises no question. Each file that already exists raises its own decision,
    with the choices that location offers. A Copy already merged there.
  - A source folder that still holds a skipped file is kept, and the item ends partly skipped, not
    as an error. A source folder the merge emptied is removed through its exact binding.
  - A folder onto a file of that name raises the type-mismatch decision (Keep both, Skip, Cancel).
    A destination folder that is a link or another reparse point, and a linked folder above it, are
    still refused before anything changes.
  - New self-test `RunFileOperationIdentityLessFolderMergeSelfTest` in
    `FileOps_ReviewRemediationContracts` (Copy merge, Move merge with a skipped child, Move merge
    that empties the source folder, link destination refused).
