# Full Debug follow-up

Fresh Full Debug at `f12acdba` passes 19 of 20 entries: 2,115 cases pass,
one fails and 52 skip. Commands passes 898/0/2, Compare 242/0/30 and
File Operations 180/0/20. Both interactive viewer entries and performance
entries pass. The prior menu-hover and ViewerImgRaw failures pass in Full.

The only failed entry is ViewerPE noninteractive. Its
`TestViewerWebTransactionalSaveAsAndCloseSafety` exceeds the unchanged
500 ms blocked-provider Close limit. The original output has no exact
elapsed time. All other blocked-provider destination/lifetime assertions pass.
The original failed aggregate and viewer output remain retained here.

Commit `cc7595de` changes only that assertion's diagnostic and its owning
specification: elapsed microseconds and provider emergency-timeout count.
The focused Debug target builds successfully. Ten isolated fresh processes
all pass, with closes of 27,330–144,864 microseconds and zero provider timeouts.
The complete noninteractive viewer group also passes; its Web close takes
32,177 microseconds with zero provider timeouts. Executable and plugin hashes
are recorded in the repetition receipts.

These repetitions do not explain the original failure or replace final Full
qualification. No product behavior, deadline, provider timeout or assertion
was changed. The failed Full run stopped the serial driver before Release or
ASan Debug began. Native ARM64 and performance acceptance remain separate.

`retained-files.json` records the exact original bytes of all copied receipts.
Full archives remain in the governed I19 evidence directory. No hosted CI ran.
