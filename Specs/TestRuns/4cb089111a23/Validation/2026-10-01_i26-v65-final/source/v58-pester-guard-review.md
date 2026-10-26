# V58 Pester source-guard review

The whitespace repair in `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` is meaningful. The old bounded character-window probe could fail solely because clang-format expanded the argument layout. The replacement at lines 4281–4284 matches the complete `PersistDirectChild` argument list with `\s*` at formatting boundaries, including `ResolveRoot()`, `_path`, `json`, the record-size cap, and both `FailIfExists` and `Replace` arms of the conditional policy. The current production call at `RedSalamander/FileOperationMoveBreadcrumb.cpp:528–533` satisfies that shape.

The supplied mutation record reports three formatting variants accepted and ten negative mutations rejected: missing `_path`, wrong argument order, an extra statement, absent `failIfExists`, either policy arm missing, absent `ResolveRoot()`, unconditional replace, missing size cap, and missing `json`. It is appropriately explicit that this is source-guard evidence (`nativeRuntimeEvidence: false`). The separate Pester result records 191 total, 191 passed, zero failed, zero skipped. The README and coverage-spec additions describe a whitespace-tolerant structural guard and retain native tests as the runtime authority; the tool inventory updates the test purpose and the explicit opt-in I26 fixture metadata.

One remaining guard-strength limitation: the regex searches the entire breadcrumb source text and does not require the call to be the `hr =` publication expression in its owning method or prove that the returned status is consumed. A commented or dead duplicate could satisfy the text probe after the real call regresses. Consider requiring the `hr =` prefix and exactly one matching publication call, or extracting the owning method before matching. This is a source-contract false-pass possibility, not evidence that the current production call is wrong.

For the new root patch proof, bind the following distinctly from the existing PluginConfig native-source guard proof:

- Root source-review `changes`: all eight exact V57→V58 paths with each `previous` and `current` SHA; exact `logicalChanges=8`, `physicalChanges=8`, manifest/inventory hashes, 85 logical/84 physical counts, and V57 equivalence path plus SHA.
- Pester guard-review artifact: schema `i26-v58-pester-source-guard-review.v1`; exact candidate test path and SHA; mutation artifact path/SHA and its source/test identities; three format variants; all ten named mutations rejected; `nativeRuntimeEvidence=false`.
- Pester execution artifact: exact `pester-results.json` path/SHA, `TotalCount=191`, `PassedCount=191`, `FailedCount=0`, `SkippedCount=0`, and provenance binding that run to the reviewed source snapshot or invocation. The JSON itself contains counts, not a source snapshot or run ID.

Observed current artifact hashes: mutation record `79f54f17453fd915d88d4cdf7b95d1e80f768778668261fbfd87750705bf8c5a`; Pester results `a58d36b63b0f2ed61e71dd3dffbcabae3f592fef237c8f1f536e7d63f74b5e89`; candidate source-contract test `d745d67a7efb4e197427f67b62cdc12419c339ff45e7a9380edb0f24a840a627`. These should be pinned by the root proof once its final schema and frozen inventory are created. ViewerPE diagnosis remains pending, so this review establishes neither V58 acceptance nor a green final Full result.


