

## V58 source-guard follow-up

Reviewed the primary `Commands.SelfTest.PluginConfig.cpp` against the manifest-backed V57 frozen copy. The frozen SHA matches the V57 manifest (`72ad7f9d73320ebfc878645296a2067dcd9813c3ae597bdd690ce6f6244a9c5c`); the primary SHA is `5b36530a9fb6109d96e4d8ae4186708eb582366a3ae2d0b8c1e2c14b4f5bdc71`. The diff is confined to the two reported whitespace-sensitive guards and explanatory comments.

**No actionable guard-strength or ownership/lifetime finding.** The Scheduled Rename guard still scopes both moved hook assignments to the extracted `AdmitScheduledRename` body. In production those assignments occur before insertion into `_tasks` and before the worker thread starts, so the `std::move` checks preserve the ownership handoff. For the mapping/path assertions, whitespace compaction preserves each original required and forbidden predicate. The required patterns correspond to the live `mappingInDestination` containment check, typed-path parent derivation, `DestinationInsideSource` rejection, and the shared `TryResolveTransferDestinationProviderPath` consumer; the three local path-guess patterns remain forbidden. The source checks remain textual assertions with their existing global-scope limitation; the edit introduces no new runtime behavior or materially weaker predicate.

The V57 frozen and current `FolderWindow.FileOperations.cpp` bytes are identical (SHA-256 `c625a03dca3195fd67682e84817f4adfc60c0fbef0bc83770031b920a60043ed`). A whitespace-insensitive diff is empty, and all 82 extracted string/character literal tokens match. Thus there is no production formatting delta to alter semantics in this source pair.

This is a source review only. The retained V57 Full failure remains a failure; this follow-up makes no green, adoption, or acceptance claim.



## Additional Phase 0 guard review

Reviewed the second primary test-source delta against the V57 frozen copy and `v58-all-pluginconfig-format-literal-audit.json`. The delta adds one whitespace-compacted view of `executorSource` and routes exactly three formerly spacing-sensitive required needles through it: native-only move mode, operation-control-cookie forwarding, and `scheduler.StartJob(this,`.

**No actionable finding.** Each new needle retains the complete significant token/value sequence, so removing the field assignment, changing either side of the cookie assignment, changing `FILESYSTEM_MOVE_NATIVE_ONLY`, or changing the scheduled owner argument from `this` causes its required `.find()` predicate to fail. The original surrounding required and forbidden checks, their Boolean conjunctions, and failure messages are unchanged. Normalization removes only spaces, tabs, CR, and LF, which resolves alignment and line-wrap drift while preserving those exact values.

The production `FolderWindow.FileOperations.State.cpp` is byte-identical to its V57 frozen manifest row (SHA-256 `0e5cc470055106ea29871022b0c1381d1292e8b25f4b19b3108300415a011a2e`). Its current code assigns `FILESYSTEM_MOVE_NATIVE_ONLY`, forwards `operationControlCookie` into the options field, and starts the per-item job with `this` as its task owner (the call is line-wrapped after `StartJob(`, which explains the false negative). This source-only check confirms the guard strings against source; it does not establish a runtime guard pass.

