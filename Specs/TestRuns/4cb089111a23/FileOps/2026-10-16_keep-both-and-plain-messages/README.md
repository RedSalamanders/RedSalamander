# Keep Both ordering and clear File Operations failures

Native x64 Debug on machine 4cb089111a23, 2026-10-16 UTC / 2026-10-17 Europe/Paris.
Branch: ej/move-existing-file-message. Base: a48e7d7880cb37ea6898d462ae188a6213a5f3f8.
The saved zero-context source-diff.txt records the final worktree changes; the per-run
workspace-final.json files retain the exact source identity of each earlier attempt.

## Behavior and coverage

The parallel Keep Both fixture waits for the nested collision's real pause handshake,
withholds a selected-root repeat grant from descendants, releases the callback pause,
and checks sibling publication/source disposition after completion. Only its serial
control requires the earlier sibling to have finished at the pause. Provider callback
serialization and product transfer concurrency are unchanged. The order metric records
concurrency in value0 and whether that sibling was already published in value1; durationUs=0
is an observation, not a timing measurement.

The existing-destination Move regression covers differing bytes, identical bytes, unsupported
destination identity for a whole root, an absent destination, and unsupported source identity.
It checks actual guard/executor/finalizer/completion behavior, byte preservation, no mutation
callbacks, no source removal, original HRESULTs and live French/English notes and Issues.

The message matrix covers 31 status scenarios, ten identity-category scenarios, and fifteen
production completed-card snapshots in both French and English. It checks plain questions,
announcements, typed and legacy Issues rows, explicit resource priority, cached uncommon system
text, Technical details and live-language completion text. Resource strings are supplied for
English, French, Czech, Slovak and Japanese. A write-protected location is distinguished from
an item's read-only attribute. The authoritative contracts are in Specs/UI/UI_FileOperationsPopup.md,
Specs/FileSystem/FileSystem_FileOperations.md, Specs/Core/Core_SharedHelpers.md and
Specs/Testing/Testing_SelfTests.md.

Broad execution also exposed three fixture defects: a directory search at the legacy path limit,
an ACL snapshot taken before Windows converted its inheritance model, and a known-size Stopping
snapshot missing its size-known fact. The repairs reuse the extended-path helper, establish the
private fixture's existing DACL/protection before taking its immutable snapshot, and retain
separate known/unknown-size controls. Exact security descriptor/control, file bytes/hash, foreground,
real RequestCancel and final result assertions remain enforced. See Microsoft's documented
[inheritance conversion](https://learn.microsoft.com/en-us/windows/win32/secauthz/automatic-propagation-of-inheritable-aces).

## Preserved attempts

These are distinct source/artifact identities, not interchangeable final evidence.

| Directory | Build receipt prefix | Passed / failed / skipped | Observation |
|---|---|---|---|
| debug-conflicts | 77f3388f07a0 | 33 / 0 / 0 | Initial ordering repair; each requested skipped case ran three times |
| debug-message-text-initial | 77f3388f07a0 | 3 / 0 / 0 | Expanded matrix and live completed-card checks before producer-check relocation |
| debug-safety-initial-failure | 77f3388f07a0 | 2 / 1 / 5 | Empty-folder inventory and exact ACL restoration failed inside the safety phase |
| debug-safety-diagnostics-failure | 6008763bd528 | 2 / 1 / 5 | Reproduced inventory failure and control 0x8004 -> 0x8404; ACL/SD lengths unchanged |
| plain-messages-debug | bcb764c0df88 | 6 / 3 / 15 | Storage repairs passed three times; stale Stopping fixture blocked later phases |
| plain-messages-final-debug | 8ae4bcf2f9cb | 9 / 3 / 12 | Safety passed three times; old OS-language assertion in producer-message test failed |
| message-review-pass-debug | 23d81b46d6b2 | 8 / 0 / 0 | Complete safety/message family passed after updating that assertion and the legacy gallery expectation |

The earlier failing parallel case and its original six skipped cases remain in the sibling
2026-10-16_move-existing-message archive. No failed attempt above is relabeled as passed.

## Closing validation

Final full-solution test-enabled x64 Debug build: 0 warnings, 0 errors. Receipt:
8f19e1856a6488bda35098c0c3a606ffa9db595a379cc37062046c13fd3d7c3d.
See build-closeout-debug.txt. Both closing runs verify and consume that receipt.

| Directory | Selection | Passed / failed / skipped |
|---|---|---|
| closeout-safety-debug | FileOpsFamily_SafetyRemediationContracts | 8 / 0 / 0 |
| closeout-conflicts-debug | FileOpsFamily_Phase09_ConflictPrompt, repeat 3 | 33 / 0 / 0 |

closing-verification.json independently checks that each of the six originally skipped cases
has exactly three passed repetitions (repeat indexes 1, 2, 3). It also retains all twelve
Keep Both scheduling observations: each parallel Copy reached its nested collision before
its sibling committed, while each parallel Move observed its sibling already committed.
All four Copy/Move x serial/parallel scenarios passed their completion checks in every repetition.
These are 41 focused native case executions, not a Full repository verdict.

pester-closeout-results.json: 214 passed, 0 failed, 0 skipped for ResourceLocalizationContracts
and TestHarnessSourceContracts. Pinned clang-format --dry-run --Werror and git diff --check pass.
The fresh closing whole-inventory check passed for 13,844 files. Explicit pre-stage validation
also passed for all 81 files in this curated run.

The final formatter samples are 1744 microseconds for French and 1509 microseconds for English,
512 mixed-status calls per language. They describe this bounded candidate workload only.
All source/spec changes and focused runs are local and uncommitted.

## Evidence boundaries

Full raw metric streams are preserved beneath
Z:\RedSalamander.Perf\evidence\keep-both-and-plain-messages before another runner invocation can
remove completed run directories. Curated metric-selection.json binds each original stream by
SHA-256, byte count and row count; selected rows retain original order. The bounded formatter
batch records 512 calls per language in microseconds. These are candidate observations without
a matched baseline, throughput/speedup claim or real NAS performance qualification.

All launched FileOps runs use --selftest-no-activate. The Release build attempt was blocked
before compilation by the independently launched Release application (PID 95428); it was not
terminated. release-build-blocked.json records that observation. No Release, ASan, ARM64,
Full repository or PR qualification is claimed. No user NAS path was accessed. Existing sandbox
disk-audit findings and failed storage-proof scratch were preserved.
