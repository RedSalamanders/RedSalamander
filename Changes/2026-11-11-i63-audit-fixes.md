- Fixes from a review of the File Operations UX work. A copy with several streams could lose its copied-file
  and folder counts for the rest of the task when one stream finished while another was still copying; the
  counts now stay. Cancel on `This location has no Recycle Bin. Delete the source files permanently?` now
  cancels the task at once instead of letting another worker ask again, and Issues shows that question
  instead of `Recycle Bin delete failed.` Before a Move removes a source at a location that cannot identify
  files, the same-file check now verifies that the location really refuses a second open of a held file,
  and refuses links and folders; where it cannot tell, the source is kept. A keep-awake request that
  Windows refused for the first running task is tried again for the next one. Reordering a queued task
  can no longer overwrite the place a task took when it was told to wait for another. The scanning
  summary of a Delete card no longer leaves an empty line when the list has no scrollbar.
