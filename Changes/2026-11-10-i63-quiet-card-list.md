- The File Operations card list moves less during a transfer. A running card now keeps the rows it
  has needed instead of changing height on every frame that needs fewer: stream rows are reserved at
  their high-water mark and released one at a time after two seconds, a two-fact row that wrapped
  stays wrapped until its facts fit with room to spare for two seconds, and the parallel-transfers
  note and the waiting reason keep their row for two seconds after they end.
  - The rate and ETA text change at most twice per second; the graph is unchanged. The row of a
    completed file keeps its bar column, so its path is not re-truncated.
  - The list's scrollbar is decided for the height the window is about to have, so it no longer
    flashes when the window is about to grow to fit the new content, and it still appears at once
    when the window cannot grow enough or when something asks to scroll meanwhile (the wheel, a
    scrollbar message, revealing the focused control). A pending or running automatic resize now uses the 16 ms
    motion cadence, and the rate history advances once per 100 ms interval even on motion ticks.
  - The rules are small state machines in `RedSalamander/FileOperationsLayoutStability.h` with
    compile-time contract checks. The popup's debug layout snapshot reports the reserved stream rows
    and the scrollbar decision.
  - Not executed: the popup self-tests need the desktop. Cases that scroll or compare card heights
    right after a change may need a settle step; they are listed in the plan's implementation log.
  - Not part of this change: tabular figures, stable stream slots in the engine, an eased height
    change, scroll anchoring, card enter and leave motion, and the sequence capture.
