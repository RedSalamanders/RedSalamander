- A name-collision question now says how the two files compare. Each file's line shows its size and
  date followed by `Newer`/`Older`/`Same date` and `Larger`/`Smaller`/`Same size`; when both are the
  same, one line says `Same size and date`, and `One hour apart, same size` covers the
  daylight-saving shift of FAT volumes and some servers. Dates within two seconds are the same.
  - The engine computes the comparison once (`FileOperations::CompareConflictItems`) and publishes it
    with the decision. The task card and the question over the pane share the text
    (`BuildConflictPromptItemText`); the question over the pane now also shows each file's size and
    date, which it did not before.
  - Eight new strings in the base and the four satellite resources.
  - Not part of this change: the exact byte count in a tooltip, the size of items on provider routes
    that report none today, and the same facts on the Issues row of a skipped or failed item.
