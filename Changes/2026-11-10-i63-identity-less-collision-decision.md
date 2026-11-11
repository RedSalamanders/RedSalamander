- A Move of a file onto an existing file on a share that cannot identify files no longer fails with
  `Errors: 2` and no choice. The task now asks, before any byte is copied, with the size and date of
  both files: Replace, Keep both, Skip, Skip all or Cancel. Keep both publishes beside the existing
  file and completes the Move; Skip keeps both files.
  - Replace on such a share uses a new optional provider contract,
    `IFileSystemExpectedOccupantPublication`, implemented by the Local provider: the new file is
    copied to a stage, and one replacing rename publishes it only while the existing file still has
    the size, dates and attributes read before the decision. A changed file replaces nothing and the
    decision is asked again; a refused rename (file in use) leaves both files and removes the stage.
    No handle is held on the existing file during the copy and no backup copy of it is made. The task
    notes once that replaced items were checked by name, size and dates.
  - A route that cannot qualify (a folder source, a destination that cannot prove what it published)
    keeps the earlier refusal.
  - `FileOps_ReviewRemediationContracts` covers Skip, Keep both, Cancel, Replace, Replace read-only, a
    file changed after the decision and a file in use at the rename, each for a share-wide and a
    single-file identity gap.
