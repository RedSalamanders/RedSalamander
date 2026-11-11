- A Move collision can now be answered `Delete source`: the existing destination file is kept, nothing is
  copied, and the source goes to the Recycle Bin. It finishes a Move that was run before. The item's result
  is `Already there; source removed`, not a skip, and the task is not Partial for it.
  - Offered only for a Managed Move whose source is on the Local file system, for a regular file that is
    provably another file than the source: by the two objects' identities, or, at a destination that cannot
    identify files, by a sharing probe (hold the destination denying all sharing, open the source by name).
  - Never a permanent delete: a source whose location has no Recycle Bin is kept, the item ends as a skip
    and a Warning says why. Permanent removal after one confirmation (the rest of D-5) is not implemented.
  - Apply to all covers it only for files with the same size and date; a remembered answer is reused only
    for such pairs. It shares its repeat key with `Replace`, so the latest answer given for all wins.
  - It is published under `More...`; on such prompts `Skip all` moves there too, because the menu now
    holds two commands.
  - New action `ConflictAction::DiscardSource` (appended, existing values unchanged), task counter and
    summary count, four strings in five resource files, policy-matrix coverage and a new self-test in
    `FileOps_ReviewRemediationContracts`.
  - Not part of this change: sources on other providers, folders and links, the one-hour class, and the
    same answer for a Move that the Local provider performs as a native rename.
