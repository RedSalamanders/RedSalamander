# File Operations P2.4 traversal-limit prompt diagnostic

- Date: 2026-08-23
- Run: `20260823T204401Z-79560-a6d33e9536c2473dba8c18f29b4a0bbc`
- Command: `./Tools/Run-AllTests.ps1 -Suite FileOps -CaseFilter Phase11_BridgeMultiFolderParallelCopyInFlightLines -SkipBuild -ValidationMode Fresh`
- Build attestation: `bbab2dcb2b701fac3cbeba6912eb921b35ce2833d05fab332142fadf05388f36`
- Outcome: failed after the quantitative traversal limit was incorrectly routed to the conflict UI.

## Evidence

The wide-shallow bridge task itself completed successfully in 99.172 seconds with 48/48 files, queue high-water 16, retained-entry high-water 21, retained-path high-water 11,634 bytes, retained-metadata high-water 9,674 bytes, and no limit hit.

The follow-on deterministic depth-limit task then logged `bridge.traversal.depthLimit` with `0x800703E9` at 22:46:03.964. Thirty milliseconds later the same terminal resource status was incorrectly logged as `item.conflict.prompt`. With no valid conflict decision for a traversal ceiling, the task waited until the phase timeout canceled it at 22:48:31.

## Resolution

Traversal ceiling statuses are resource-policy stops, not conflicts. Reject them before every provider callback/outer conflict entry point and preserve the exact status. The task must finish with the prior item published, the bounded item failed, no prompt, and no continued discovery below the rejected depth.
