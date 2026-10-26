# Release validation interrupted by host restart

The first `FileOpsFamily_Phase11_BridgeAndConnections` invocation,
`20261019T053144Z-45640-25dd39acc8204cff891dcc002d1a40da`, has no final test result.
Its last native trace was written at 2026-10-19 05:32:35.674 UTC; Windows System event
1074 records a restart request at 05:32:35.845 UTC. Later restart events identify a
planned operating-system upgrade. The partial run is neither a pass nor a failure.

Before rerunning, the complete raw artifacts and original controller logs were preserved in
`Z:\RedSalamander.Perf\evidence\final-cancellation-release-2026-10-19\interrupted-controller`.
The curated [classification](../../Validation/2026-10-19_final-cancellation-release/host-restart-interruption/classification.json),
[host events](../../Validation/2026-10-19_final-cancellation-release/host-restart-interruption/host-events.json),
[partial trace](../../Validation/2026-10-19_final-cancellation-release/host-restart-interruption/trace.txt)
and digest-bound selected metrics remain in the same Release qualification packet.

The artifact profile lock was cleared cleanly. The unchanged source snapshot, test-enabled
Release receipt and every attested output hash were verified before continuing only the
unfinished selections. The original four successful selections retain their original receipts.
This is selected local continuation evidence, not a Fresh Full or repository-wide Resume verdict.

Reproduce the interrupted selection from its exact receipt-bound isolated checkout with
`Tools/Run-AllTests.ps1 -Suite FileOps -CaseFilter FileOpsFamily_Phase11_BridgeAndConnections
-Configuration Release -Platform x64 -SkipBuild -TestRoot Z:/RedSalamander.Perf
-AllowAlternateVolumeTestRoot` and a new run ID. See the
[Release packet](../../Validation/2026-10-19_final-cancellation-release/README.md)
for actual final results, source identity, scope and limitations.
