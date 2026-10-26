# History-popup diagnostic source guard

The `2f4fb03f` Debug Full general tooling entry reports 774 passed and one failed.
Its source guard expected the popup query inline in the success condition; the
diagnostic implementation stores the same result before testing it with the keyboard
selection. This is a source-form mismatch, not a weakened native assertion.

The focused original policy reproduces **187 passed, one failed**. The corrected
source-contract file plus ToolInventory and ToolingGovernance tests pass **217/217**.
Both inventories pass. The new guard requires the query assignment, the bounded
three-second loop and capture only when the query succeeds and keyboard selection exists.
The immediate-query-failure rejection now also covers the canonical namespace.

Before/after native fixture SHA-256 is identical: `1AED7522DFC59EDD33F303076D8835D8FBFA6228846B4D2E18EDFB67B9AC55D5`.
The JSON receipts bind each measured policy by SHA-256. Their `source` is the checked-out
documentation HEAD; the after policy was then an uncommitted candidate. This record
does not replace Fresh Full or prove that the earlier native popup failure is fixed.
No native UI run or hosted CI is launched by these focused source-policy checks.

The six original receipts are hashed in `retained-files.json` and retained at
`Z:/RedSalamander.Perf/evidence/I19-20260909/history-source-contract`. [I19](../../../../Plans/Done/DxUi_SharedLibraryAdoptionAndReleasePlan_2026-09-09.md)
owns final qualification; ARM64/ASan remain user-deferred.
