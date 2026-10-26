# DxUi PR #69 final hosted CI evidence

This archive records DxUi validation for PR #69 at head `d404ec1b0f656818338d868a30b7e8ece644e8ab`. That head only corrected a missing trailing `F` in a documentation fingerprint; the DxUi library source and the paired candidate input fingerprint were unchanged. The paired comparisons use baseline merge parent `bea676a1f841` and candidate merge ref `4c2aefd130ff` (`refs/pull/69/merge`). The candidate library-input fingerprint is `F826386B05B760C3E8FFFFEEB857B3F889398C1E10CD57AEBBB048D20AD5E54F`; baseline is `41423433DDF45EFBBBAFA7B35D9A16D2F9D057E4271779D4D8EAC4955C481051`.

`manifest.txt` lists SHA-256 and byte length for each copied report/log plus `artifact-receipts.txt`. The copied report/log bytes were taken from the listed GitHub artifact entries. Archive-local `.gitattributes` disables text normalization. Each archived evidence file is below 5 MiB; no capture PNGs or full artifact ZIPs are included here. Original ZIPs remain in the marked `D:\RedSalamander.Perf\evidence\dxui-ci-d404ec1-37862075647` source-evidence tree; `artifact-receipts.txt` records their artifact IDs and SHA-256 values.

## Run outcomes

Workflow run `37862075647` was invoked by the `d404ec1` head. Attempt 1 ended with an x64 Release Menu test failure and an inconclusive paired gate. The five other first-attempt profiles passed: x64 Debug, x64 ASan Debug, ARM64 Debug, ARM64 Release, and ARM64 ASan Debug. The failure was `TestMenuHoveringSiblingClosesOpenSubmenuAfterDelay`: the log reports `FAILED: hovering B2 updates the root-popup hover target before delayed submenu close` and the Menu result report has exit code 1.

Attempt 2 reran only failed x64 Release job `113609797975`. It passed; the same Menu case completed successfully and its Menu result report has exit code 0. This one pass does not establish why attempt 1 failed, so the cause remains unresolved. The earlier five native-profile results were not rerun in this attempt.

Attempt 3 reran only the paired benchmark job `113614291427`. It ended Inconclusive. All six native profile statuses in the final workflow attempt were successful, including the x64 Release retry and hosted ARM64 DxUi library profiles. The overall workflow remained failed because the paired performance gate was inconclusive. No further retry was performed.

## Paired results and residual gate

Both the original paired attempt and its single permitted retry used x64 Release on a shared hosted VM, with three repetitions of interleaved A, B, B, A passes (six runs per side), exact two-sided Mann–Whitney U testing, a 5% timing band and 2% memory band. Neither result supports a performance-gain claim, and neither confirms a source regression.

The first paired attempt was Inconclusive. Its flagged MultilineGrid dirty `composeCpuP95Ms` changed from 0.0166 ms to 0.0201 ms (+21.45%, p=0.0281 against a 5% band), while same-binary controls drifted 22.4%. Default and MultilineGridDistinct were within-noise-budget.

The single paired retry was also Inconclusive. Default clean `fps` changed from 277.1 to 244.5 (-11.76%, p=0.0411 against a 5% band), while controls drifted 13.3%. MultilineGridDistinct clean `prepareP95Ms` changed from 0.0006 ms to 0.0007 ms (+16.67%, p=0.0152 against a 5% band), while controls drifted 14.3%. MultilineGrid was within-noise-budget. Controls were stable in 0/6 same-binary comparisons for every scenario in both attempts, leaving the flagged values unresolved as source effects versus runner noise. The gate residue is therefore an unresolved performance comparison; there is no further retry authorized by this campaign.

The `native (ARM64, windows-11-vs2026-arm, ...)` results qualify the DxUi library's hosted ARM64 profile only. They are not RedSalamander application ARM64 qualification; no app ARM64 result is included here.

## Curated files

- `native-x64-release-attempt1/` and `native-x64-release-attempt2/` contain the exact Menu test log and suite report for the failure and its successful rerun.
- `paired-attempt1-inconclusive/` and `paired-attempt3-final-retry-inconclusive/` contain the exact Markdown verdict, structured comparison verdict, and report summary from each paired attempt. Structured JSON filenames use `.comparison.json` to identify comparison records.
- `artifact-receipts.txt` identifies the originating workflow jobs and raw GitHub artifact ZIP hashes. `manifest.txt` hashes each copied evidence output and the receipt.