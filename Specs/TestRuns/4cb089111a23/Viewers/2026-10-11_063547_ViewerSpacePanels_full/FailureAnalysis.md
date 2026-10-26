# Fresh Full Release outcome

Run `20261011T040752Z-65208-c8d6028187b24b18b10e1ff93622171b` completed all 20 entries in 52m 53.5s. Repository verdict: **FAILED**. Fifteen entries promoted. The legacy report records 2,136 passed, 13 failed and 276 skipped cases. The build completed without warnings or errors; its receipt is archived. Both interactive viewer entries passed.

The raw v1 report preserves every declared case status and skip reason; v2 preserves gate, source snapshot, build receipt and promotion identities. `full-report-bindings.json` records original byte counts and SHA256 for these unmodified reports.

Five entries failed:

- ViewerPETests.Noninteractive: ViewerPE's blocked-read Close assertion measured 551 ms against 500 ms. All seven ViewerSpace cases in this entry passed.
- PluginContractTests: the S3 multipart export reported 49 passed and one failed check. This report does not establish which check failed. Other exported contract groups passed; documented Release-only omissions remain in their raw output.
- ToolsPesterTests: 928 passed and one failed. The protected Done-history contract finds two unexpected, already committed documents from `2ae984373`: `CodeReview_MasterMerges_2026-10-05/findings-ledger.md` and `CodeReview_MasterMerges_Delivered_2026-10-10.md`. It found no modified or missing baseline document. ViewerSpace did not change these paths.
- Commands: 936 passed, eight failed, six skipped. The ViewerSpace capture check expects background at the bottom-left edge of the analysis child, now occupied by the right-panel splitter. Both text witnesses passed. The subsequent test correction checks the canonical divider color at the affected gutter while preserving all eight composition witnesses. That filtered rerun passed the five theme frames, then exposed the old 456-DIP height expectation; the new separator requires the already specified 462-DIP minimum. Both threshold witnesses are corrected in a later source. This archive precedes both corrections. Another source guard requires exact single-space assignment formatting although the clipboard-admission assignment is column-aligned. The remaining failures concern Preferences focus restoration, Find quiescence, Compare options body composition, File Operations path tooltip and disclosure geometry, and thumbnail slider input. See the unmodified case report for exact reasons.
- FileOperations: two passed, two failed, 240 skipped after the suite's mandatory safety prerequisites failed. The first failure concerns a missing refusal diagnostic for `trailing-dot.`; the second is the safety-remediation contract. These skips are not accepted qualification.

No failing assertion has been weakened, and no unrelated production code has been changed to make this Full result pass. Later runs must be assessed against their own source and artifact identities.
