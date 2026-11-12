- During a parallel copy or move, each file now stays on its own line of the card while it is
  copied: when one file finishes, the files below it no longer move up a line. The numbers in the
  progress, counts and speed rows keep their width as they change.
  - The engine gives each transfer stream a row slot for its lifetime (the lowest free slot) and
    keeps it when the in-flight array is compacted; the popup snapshot carries the slot and the
    card draws each stream in its slot. A finished file's row stays for 300 ms, then the row is
    blank until another file takes it or the card's reservation releases it.
  - The progress, counts and ETA, and throughput rows use tabular figures on the popup's own text
    layouts. The skipped-counting work summary is measured once per text, width and format
    instead of on every paint.
  - New ETW counters `FileOps.Popup.CardHeightChanges`, `FileOps.Popup.ScrollBarToggles` and
    `FileOps.Popup.WindowResizes`, and new self-test `RunFileOperationStreamSlotSelfTest` in
    `FileOps_ReviewRemediationContracts`.
  - Not part of this change: a reserved width for the widest form of each number, the eased
    height change, scroll anchoring, card enter and leave motion, and the budget case
    `FileOps_PopupLayoutStability`.
