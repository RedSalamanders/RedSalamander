# I26 V48 source-delta review addendum

Root reviewed the candidate against frozen V47, with read-only Sol/Luna reviews and independent source/API checks. This document is source review, not native qualification or performance acceptance.

V48 manifest SHA-256: 5ef0aab33a628e8e0dc841957bb7307c7e59b2d2b00cb6b644c5ab46b4c08c3f. Root checked all 82 logical/81 physical paths against both candidate checkout and frozen source; no hash mismatch. Consumer HEAD remains ccebd5af68eb45f2e9deed64a7063780e71eb1ce; qualified DxUi pin 49c963f5751a6754e6b08e0362b96cfc7fc64414 unchanged.

## Product delta

Full-path read-only detail now uses existing pinned TextField capabilities within the same popup host. Four child controls are created on demand, retired on close/hide or row removal, and released safely outside their callback. Normal task subtrees and decision shortcuts are unavailable while detail is active. Focus/selection are preserved for an unchanged value across rebuild. Long Unicode values support full or selected copy and actual viewport movement. No DxUi change, additional HWND, timer, worker or plugin ABI ownership.

See path-detail-source-review-v48.md for API/lifecycle review corrections and the gallery matrix. All new IDs have EN/FR/cs-CZ/ja-JP/sk-SK parity; resource contracts 9/0/0. Existing V47 captures are pre-feature baseline evidence, not full-path-detail qualification.

## Validation delta

Commands adds >4Ki French UNC/combining/Hebrew/emoji, exact UIA TextPattern, selection/copy, keyboard viewport, read-only input, stale-provider retirement, root replacement, hide/removal and 40-cycle resource witnesses. Gallery adds root/scrolled dark/light on each monitor and actual same-HWND DPI hop without native activation. Native runs pending.

V47 Full reported 2204 passed / 6 failed / 46 skipped, clean sandbox. Reviewed fixture corrections and unresolved S3/hidden-current diagnostics are in v48-commands-repair-source-review.md. Candidate additionally binds ViewCommands explicitly. No existing failures are waived. Static harness source contracts 191/0/0 in v48-harness-source-contracts-r3.xml (SHA f074eb1b2e246f53e3074b812eb043866e5dd07ab00dd723dc54af8c97305094).

## Performance boundary

A78/B49 paired fixture remains byte-identical to approved 90f7301cc91e96d50985bf61451d6dcb1799254a5de6212928978137757ce740. New detail is closed in that fixture, so the explicit open/close Commands memory samples remain separate evidence. Final pairing must bind the ultimately qualified post-Full candidate receipt, not V47 or this unvalidated build. No timing/resource acceptance is inferred from static review.
